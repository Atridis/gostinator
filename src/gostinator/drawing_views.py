"""Create and place the necessary native orthographic and section views."""
from dataclasses import dataclass
from math import dist, isfinite, pi

from pythoncom import VT_DISPATCH, com_error
from win32com.client import VARIANT

from .model_geometry import VIEWS, choose_section
from .solidworks import as_dispatch, get_value


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
        if get_value(doc, "MakeSectionLine") is False:
            return None
        if not get_value(segment, "Select4", False, null_object):
            return None
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
            _create(doc, part, slot.axis, slot.sign, position, scale)
    get_value(doc, "EditRebuild3")
    views = drawing_views(doc)
    if not views:
        raise RuntimeError("После построения не найдено связанных видов")
    return [str(get_value(view, "GetName2")) for view in views]


def create_flat_pattern(doc, part, layout):
    slot = layout.slots[-1]
    view = get_value(doc, "CreateFlatPatternViewFromModelView3",
                     str(part), "", *layout.positions[-1], 0., False, False)
    if view is None or not get_value(view, "IsFlatPatternView"):
        raise RuntimeError("SolidWorks не подтвердил нативную развёртку")
    view.UseSheetScale = False
    view.ScaleDecimal = layout.scale * slot.factor
    view.ShowSheetMetalBendNotes = False
    get_value(doc, "EditRebuild3")
    return str(get_value(view, "GetName2"))
