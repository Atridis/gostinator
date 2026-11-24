"""Create and place the necessary native orthographic and section views."""
from dataclasses import dataclass
from math import dist, isfinite, pi

from pythoncom import VT_DISPATCH, com_error
from win32com.client import VARIANT

from .model_geometry import VIEWS, choose_section, read_model
from .solidworks import as_dispatch, get_value


SW_LINE_CENTER = 4
SW_LINE_WEIGHT_THIN = 0
SW_UNSUPPRESS_FEATURE = 1
SW_THIS_CONFIGURATION = 1


@dataclass(frozen=True)
class Frame:
    width: float
    height: float
    first_angle: bool

    @property
    def drawable(self):
        return (.020, .060, self.width - .005, self.height - .005)


def sheet_frame(doc):
    sheet = get_value(doc, "GetCurrentSheet")
    properties = get_value(sheet, "GetProperties2")
    return Frame(float(properties[5]), float(properties[6]), bool(properties[4]))


def drawing_views(doc):
    return [as_dispatch(view) for sheet in (get_value(doc, "GetViews") or ()) for view in sheet[1:]]


def transform_of(view):
    values = get_value(get_value(view, "ModelToViewTransform"), "ArrayData")
    if not values or len(values) != 16 or not all(map(isfinite, values)) or values[12] <= 0:
        raise RuntimeError("Вид без корректного преобразования координат")
    return values


def project(point, matrix):
    return tuple(sum(point[i] * matrix[3*i + j] for i in range(3)) * matrix[12] + matrix[9 + j]
                 for j in range(3))


def add_view_centerlines(doc, view, axes, *, circles=()):
    if not axes and not circles:
        return 0
    box = get_value(view, "GetOutline")[:4]
    axes = [line for line in axes if all(box[i] - .002 <= point[i] <= box[i+2] + .002
                                       for point in line for i in (0, 1))]
    circles = [(center, radius) for center, radius in circles
               if all(box[i] <= center[i]-radius and center[i]+radius <= box[i+2] for i in (0, 1))]
    if not axes and not circles:
        return 0
    get_value(doc, "ClearSelection2", True)
    if not get_value(doc, "ActivateView", get_value(view, "GetName2")):
        return 0
    try:
        sketch = get_value(view, "GetSketch")
        transform = get_value(sketch, "ModelToSketchTransform")
        direct = get_value(transform, "ArrayData")
        inverse = get_value(get_value(transform, "Inverse"), "ArrayData")
        manager = get_value(doc, "SketchManager")
        added = 0
        def line_segment(a, b):
            local = tuple(project(point, direct) for point in (a, b))
            if any(dist(project(point, inverse), original) > 1e-7
                   for point, original in zip(local, (a, b))):
                return None
            return get_value(manager, "CreateLine", *local[0], *local[1])

        segments = (line_segment(a, b) for a, b in axes)
        circular = (get_value(manager, "CreateCircleByRadius", *project(center, direct), radius*direct[12])
                    for center, radius in circles)
        for segment in (*segments, *circular):
            if segment is not None:
                segment.Style = SW_LINE_CENTER
                segment.Width = SW_LINE_WEIGHT_THIN
                segment.Color = 0
                added += 1
        return added
    finally:
        get_value(doc, "ClearSelection2", True)
        get_value(doc, "ActivateView", "")


def _create(doc, part, axis, sign, position, scale):
    name = VIEWS[axis, sign]
    view = get_value(doc, "CreateDrawViewFromModelView3", str(part), name, *position, 0.)
    if view is None:
        local_names = {"*Front": "*Спереди", "*Back": "*Сзади",
                       "*Left": "*Слева", "*Right": "*Справа",
                       "*Top": "*Сверху", "*Bottom": "*Снизу"}
        view = get_value(doc, "CreateDrawViewFromModelView3", str(part), local_names[name], *position, 0.)
    if view is None:
        raise RuntimeError(f"SolidWorks не создал вид {name}")
    view.UseSheetScale = False
    view.ScaleDecimal = scale
    return view


def _delete_selected(doc):
    return get_value(get_value(doc, "Extension"), "DeleteSelection2", 0)


def _create_section(doc, front, model, axis, section, destination, scale, *, first_angle, turned):
    matrix = transform_of(front)
    x0, y0, x1, y1 = get_value(front, "GetOutline")[:4]
    model_point = list(section.bores[0].point((model.bounds[axis] + model.bounds[axis+3]) / 2))
    cross = next(i for i in range(3) if i != axis and i != section.along)
    model_point[cross] = section.coordinate
    x, y = project(model_point, matrix)[:2]
    horizontal = abs(matrix[3 * section.along]) > .9
    pad = .003
    endpoints = ((x0 - pad, y, 0.), (x1 + pad, y, 0.)) if horizontal else (
        (x, y0 - pad, 0.), (x, y1 + pad, 0.))
    get_value(doc, "ClearSelection2", True)
    name = get_value(front, "GetName2")
    if not get_value(doc, "ActivateView", name):
        return None
    sketch = get_value(front, "GetSketch")
    transform = get_value(sketch, "ModelToSketchTransform")
    direct = get_value(transform, "ArrayData")
    inverse = get_value(get_value(transform, "Inverse"), "ArrayData")
    local = [project(point, direct) for point in endpoints]
    if any(dist(project(point, inverse), original) > 1e-7
           for point, original in zip(local, endpoints)):
        return None
    segment = get_value(get_value(doc, "SketchManager"), "CreateLine", *local[0], *local[1])
    if segment is None:
        return None
    try:
        null_object = VARIANT(VT_DISPATCH, None)
        if not get_value(segment, "Select4", False, null_object):
            return None
        # CreateSectionViewAt5 promotes the selected sketch segment itself.
        # A drawing of a part has no excluded components.
        # Orient a horizontal cut to face the first-angle placement. For a
        # single central through bore, either section direction cuts the
        # same geometry; show the side facing the other projection.
        coaxial = len({b.center for b in section.bores}) == 1
        through = any(b.high - b.low >= .85 * model.spans[axis] for b in section.bores)
        flip = ((horizontal and first_angle) or (not horizontal and coaxial and through)) ^ turned
        view = get_value(doc, "CreateSectionViewAt5", *destination, 0., "A", 4 if flip else 0, "", 0.)
        if view is not None and get_value(view, "Type") == 2:
            view.UseSheetScale = False
            view.ScaleDecimal = scale
            segment = None  # promoted to the native section line
            return view
        return None
    finally:
        if segment is not None:
            get_value(doc, "ClearSelection2", True)
            if get_value(segment, "Select4", False, VARIANT(VT_DISPATCH, None)):
                _delete_selected(doc)
        get_value(doc, "ClearSelection2", True)


def create_views(doc, part, model, frame, layout=None):
    if model is None:
        method = "Create1stAngleViews2" if frame.first_angle else "Create3rdAngleViews2"
        if not get_value(doc, method, str(part)):
            raise RuntimeError("SolidWorks не создал стандартные виды")
        return [str(get_value(v, "GetName2")) for v in drawing_views(doc)]

    primary = layout.slots[0]
    front = _create(doc, part, primary.axis, primary.sign, layout.positions[0], layout.scale)
    if primary.turn:
        front.Angle = pi / 2
    get_value(doc, "EditRebuild3")
    for slot, position in zip(layout.slots[1:], layout.positions[1:]):
        if slot.role == "flat":
            continue
        scale = layout.scale * slot.factor
        if slot.role == "section":
            section = choose_section(model, primary.axis)
            try:
                cut = _create_section(doc, front, model, primary.axis, section, position, scale,
                                      first_angle=frame.first_angle, turned=primary.turn)
            except (AttributeError, RuntimeError, com_error) as exc:
                print(f"[WARN] Разрез пропущен — {exc}", flush=True)
                cut = None
            if cut is None:
                _create(doc, part, slot.axis, slot.sign, position, scale)
        else:
            view = _create(doc, part, slot.axis, slot.sign, position, scale)
            if slot.turn:
                view.Angle = pi / 2
    get_value(doc, "EditRebuild3")
    views = drawing_views(doc)
    if not views:
        raise RuntimeError("После построения не найдено связанных видов")
    return [str(get_value(view, "GetName2")) for view in views]


def _flat_plane(part, thickness):
    bodies = get_value(part, "GetBodies2", 0, False) or ()
    if len(bodies) != 1:
        return None
    body = as_dispatch(bodies[0])
    faces = (as_dispatch(raw) for raw in get_value(body, "GetFaces") or ())
    planes = [(get_value(face, "GetArea"), normal) for face in faces
              if (normal := get_value(face, "Normal")) and sum(v*v for v in normal) > .99]
    if not planes:
        return None
    _, normal = max(planes, key=lambda row: row[0])
    points = (get_value(as_dispatch(vertex), "GetPoint")
              for vertex in get_value(body, "GetVertices") or ())
    depths = [sum(a*b for a, b in zip(point, normal)) for point in points if point]
    if depths and abs(max(depths) - min(depths) - thickness) <= max(.000075, .005*thickness):
        return normal
    return None


def _prepare_flat_configuration(part, folded_name, flat_name):
    if not flat_name or flat_name == folded_name or not get_value(part, "ShowConfiguration2", flat_name):
        raise RuntimeError("SolidWorks не переключился на отдельную конфигурацию развёртки")
    try:
        patterns = []
        thickness = None
        feature = get_value(part, "FirstFeature")
        while feature is not None:
            kind = get_value(feature, "GetTypeName2")
            if kind == "FlatPattern":
                patterns.append(feature)
            elif kind in ("SheetMetal", "SMBaseFlange") and thickness is None:
                definition = get_value(feature, "GetDefinition")
                candidate = get_value(definition, "Thickness") if definition is not None else None
                if isinstance(candidate, (float, int)) and candidate > 0:
                    thickness = candidate
            feature = get_value(feature, "GetNextFeature")
        if thickness is None:
            raise RuntimeError("В модели не найдена толщина листового металла")
        if len(patterns) == 1 and any(get_value(patterns[0], "IsSuppressed2", SW_THIS_CONFIGURATION, None) or ()):
            if not get_value(patterns[0], "SetSuppression2", SW_UNSUPPRESS_FEATURE, SW_THIS_CONFIGURATION, None):
                raise RuntimeError("SolidWorks не включил операцию разворачивания листа")
            if not get_value(part, "ForceRebuild3", False):
                raise RuntimeError("SolidWorks не перестроил конфигурацию развёртки")
        normal = _flat_plane(part, thickness)
        if normal is None:
            raise RuntimeError("Конфигурация развёртки не содержит плоской заготовки")
        try:
            blank = read_model(part)
        except (AttributeError, com_error):
            blank = None
        return normal, blank
    finally:
        if not get_value(part, "ShowConfiguration2", folded_name):
            raise RuntimeError("SolidWorks не вернул согнутую конфигурацию модели")


def _flat_view_faces_blank(view, plane):
    normal = transform_of(view)[2:9:3]
    return abs(sum(a*b for a, b in zip(normal, plane))) >= .999


def create_flat_pattern(doc, part, layout, writable_part):
    from .annotation_layout import FLAT_PATTERN_CAPTION, _place_flat_caption, geometry_outline, view_strokes

    slot = layout.slots[-1]
    manager = get_value(writable_part, "ConfigurationManager")
    folded_name = get_value(get_value(manager, "ActiveConfiguration"), "Name")

    def insert(config):
        view = get_value(doc, "CreateFlatPatternViewFromModelView3",
                         str(part), config, *layout.positions[-1], 0., False, False)
        if view is None or not get_value(view, "IsFlatPatternView"):
            raise RuntimeError("SolidWorks не подтвердил нативную развёртку")
        return view

    view = insert(folded_name)
    flat_name = get_value(view, "ReferencedConfiguration")
    plane, blank = _prepare_flat_configuration(writable_part, folded_name, flat_name)
    get_value(doc, "EditRebuild3")
    if not _flat_view_faces_blank(view, plane):
        get_value(doc, "ClearSelection2", True)
        extension = get_value(doc, "Extension")
        if not (get_value(extension, "SelectByID2", get_value(view, "GetName2"), "DRAWINGVIEW",
                          0., 0., 0., False, 0, VARIANT(VT_DISPATCH, None), 0)
                and _delete_selected(doc)):
            raise RuntimeError("SolidWorks не заменил торцевой вид развёртки")
        view = insert(folded_name)
        if get_value(view, "ReferencedConfiguration") != flat_name or not _flat_view_faces_blank(view, plane):
            raise RuntimeError("SolidWorks не построил вид нормалью к плоской заготовке")
    view.UseSheetScale = False
    view.ScaleDecimal = layout.scale * slot.factor
    view.ShowSheetMetalBendNotes = False
    get_value(doc, "EditRebuild3")
    if slot.turn:
        x0, y0, x1, y1 = get_value(view, "GetOutline")[:4]
        if x1 - x0 > y1 - y0:
            view.Angle = get_value(view, "Angle") + pi / 2
            get_value(doc, "EditRebuild3")
    name = str(get_value(view, "GetName2"))
    get_value(doc, "ClearSelection2", True)
    if get_value(doc, "ActivateView", name):
        try:
            note = get_value(doc, "InsertNote", FLAT_PATTERN_CAPTION)
            if note is not None:
                annotation = get_value(note, "GetAnnotation")
                font = get_value(annotation, "GetTextFormat", 0)
                font.CharHeight = .004
                get_value(annotation, "SetTextFormat", 0, False, font)
                get_value(doc, "ViewZoomtofit2")
                outline = get_value(view, "GetOutline")[:4]
                _place_flat_caption(annotation, geometry_outline(view_strokes(view, outline), outline))
        finally:
            get_value(doc, "ClearSelection2", True)
            get_value(doc, "ActivateView", "")
    return name, blank
