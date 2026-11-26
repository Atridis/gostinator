"""Center marks and shared diameter callouts for visible holes."""
from collections import defaultdict
from dataclasses import dataclass
from math import atan2, cos, dist, isclose, sin, tau
from re import fullmatch

from pythoncom import VT_DISPATCH, com_error
from win32com.client import VARIANT

from .annotation_layout import DimensionTextLayout, _line_crossings, dimension_lines, text_geometry
from .drawing_dimensions import (
    Circle, DIAMETER, Task, VISIBLE, _create, _discard, _hide, _optional, _plain_diameter, _projected_bore,
)
from .drawing_views import add_view_centerlines, project
from .model_geometry import Bore, POSITION_TOLERANCE as TOLERANCE, axis_of
from .solidworks import as_dispatch, get_value


@dataclass(frozen=True)
class ThreadSpec:
    size: str
    diameter: int | None
    depth: int | None
    fit: str | None
    tap_drill: int | None = None


@dataclass(frozen=True)
class Hole:
    center: tuple[float, float]
    rim: Circle
    bores: tuple
    profile: tuple
    thread: str = ""
    edges: tuple[Circle, ...] = ()
    tap_drill: Bore | None = None


def _stations(values):
    return tuple(round(value / TOLERANCE) for value in values)


def _thread_name(size):
    return size.upper().replace("М", "M").replace(",", ".")


def _thread_centers(view):
    """Use drawn cosmetic threads to distinguish taps from plain bores."""
    result = []
    part = _optional(view, "ReferencedDocument")
    thread = _optional(view, "GetFirstCThread")
    while thread is not None:
        annotation = _optional(thread, "GetAnnotation")
        if annotation is not None and _optional(annotation, "Visible") == 1:
            name = _optional(annotation, "GetName")
            feature = _optional(part, "FeatureByName", name) if part is not None and name else None
            definition = _optional(feature, "GetDefinition") if feature is not None else None
            diameter = _optional(definition, "Diameter") if definition is not None else None
            owner = _optional(feature, "GetOwnerFeature") if feature is not None else None
            wizard = _optional(owner, "GetDefinition") if owner is not None else None
            size = _optional(wizard, "FastenerSize") if wizard is not None else None
            source = _optional(definition, "ThreadCallout") if definition is not None else None
            if not size and isinstance(source, str):
                size = next((word for word in source.split() if fullmatch(r"[MМ]\d+(?:[.,]\d+)?", word)), "")
            depth = _optional(wizard, "ThreadDepth") if wizard is not None else None
            fit = _optional(wizard, "ThreadClass") if wizard is not None else None
            drill_sizes = {round(value/TOLERANCE) for prop in ("TapDrillDiameter", "ThruTapDrillDiameter")
                           if wizard is not None and isinstance(value := _optional(wizard, prop), (float, int))
                           and value > TOLERANCE}
            spec = ThreadSpec(_thread_name(size) if isinstance(size, str) else "",
                              round(diameter/TOLERANCE) if isinstance(diameter, (float, int)) else None,
                              round(depth/TOLERANCE) if isinstance(depth, (float, int)) else None, fit,
                              next(iter(drill_sizes)) if len(drill_sizes) == 1 else None)
            data = _optional(annotation, "GetDisplayData")
            count = (_optional(data, "GetArcCount") or 0) if data is not None else 0
            for index in range(count):
                arc = _optional(data, "GetArcAtIndex2", index)
                if arc and len(arc) == 17:
                    result.append((tuple(arc[10:12]), spec))
        thread = _optional(thread, "GetNext")
    return result


def projected_holes(circles, model, matrix, threads=()):
    normal = axis_of(matrix[2:9:3])
    if normal is None or model is None:
        return ()
    by_center = defaultdict(list)
    for bore in model.bores:
        if bore.axis == normal:
            by_center[_stations(bore.center)].append(bore)
    rings = defaultdict(list)
    for circle in circles:
        if circle.axis == normal:
            station = _stations(value for i, value in enumerate(circle.anchor.model) if i != normal)
            if station in by_center:
                rings[station].append(circle)

    result = []
    for station, candidates in rings.items():
        bores = by_center[station]
        matched = [c for c in candidates if any(abs(c.radius-bore.radius) <= TOLERANCE
                   and bore.low-TOLERANCE <= c.anchor.model[normal] <= bore.high+TOLERANCE for bore in bores)]
        if not matched:
            continue
        outside = max((bore.radius for bore in bores), default=0.)
        rim = max((c for c in candidates if c.fraction >= .80 and c.radius <= outside + .002),
                  key=lambda c: (c.radius, c.anchor.sheet[2]), default=max(matched, key=lambda c: c.radius))
        center = rim.anchor.sheet[:2]
        sources = [spec for placement, spec in threads if dist(placement, center) <= 4*TOLERANCE*matrix[12]]
        # Same visible opening, entry station and full depth profile. A
        # matching diameter alone cannot identify a stepped or threaded hole.
        profile = tuple(sorted({_stations((b.radius, b.low, b.high)) for b in bores}))
        visible_rings = tuple(sorted({_stations((c.radius, c.anchor.model[normal])) for c in candidates
                                      if c.fraction >= .80}))
        thread_profile = sources[0] if len(set(sources)) == 1 else None
        thread = thread_profile.size if thread_profile is not None else ""
        tap_drills = [bore for bore in bores if thread and thread_profile.tap_drill is not None
                      and abs(2*bore.radius-thread_profile.tap_drill*TOLERANCE) <= TOLERANCE]
        result.append(Hole(center, rim, tuple(bores), (profile, visible_rings, tuple(sorted(set(sources), key=str))),
                           thread, tuple(matched), tap_drills[0] if len(tap_drills) == 1 else None))
    return tuple(result)


def _groups(holes):
    groups = defaultdict(list)
    for hole in holes:
        groups[hole.profile].append(hole)
    return tuple(tuple(group) for group in groups.values() if len(group) > 1)


def _pitch_circle(group, holes, scale):
    if len(group) < 3:
        return None
    points = [hole.center for hole in group]
    center = tuple(sum(values)/len(points) for values in zip(*points))
    tolerance = 2*TOLERANCE*scale
    if not any(dist(h.center, center) <= tolerance and h.rim.radius > max(g.rim.radius for g in group)
               for h in holes):
        return None  # a square grid alone is not evidence of a bolt circle
    radius = dist(center, points[0])
    if any(abs(dist(center, p)-radius) > tolerance for p in points[1:]):
        return None
    angles = sorted(atan2(y-center[1], x-center[0]) for x, y in points)
    steps = [b-a for a, b in zip(angles, (*angles[1:], angles[0]+tau))]
    if all(abs(step-tau/len(points))*radius <= tolerance for step in steps):
        return (*center, 0.), radius
    return None


def axis_connections(holes, scale, *, groups=None):
    """Connect only gaps between the centers of a common row or column."""
    tolerance = max(TOLERANCE*scale, 1e-6)
    for group in _groups(holes) if groups is None else groups:
        for along in (0, 1):
            cross = 1-along
            rows = defaultdict(list)
            for hole in group:
                rows[round(hole.center[cross]/tolerance)].append(hole)
            for row in rows.values():
                row.sort(key=lambda h: h.center[along])
                for first, last in zip(row, row[1:]):
                    start, end = first.center[along], last.center[along]
                    offset = (first.center[cross]+last.center[cross])/2
                    if any(start+tolerance < other.center[along] < end-tolerance
                           and abs(other.center[cross]-offset) < other.rim.radius*scale
                           for other in holes if other is not first and other is not last):
                        continue
                    # The native marks already draw through the circles.
                    low = start+first.rim.radius*scale
                    high = end-last.rim.radius*scale
                    if high-low > .001:
                        yield ((low, offset, 0.), (high, offset, 0.)) if along == 0 else (
                            (offset, low, 0.), (offset, high, 0.))


def _mark(doc, view, hole, scale):
    get_value(doc, "ClearSelection2", True)
    if not _optional(view, "SelectEntity", hole.rim.anchor.entity, False):
        return False
    mark = _optional(doc, "InsertCenterMark3", 2, False, False)  # native single center mark
    if mark is None:
        return False
    mark.UseDocDisplaySettings = False
    mark.ShowLines = True
    mark.CenterLineFont = True
    mark.Size = .0015 / max(1., scale)
    return True


def add_hole_center_marks(doc, view, holes, scale):
    if not holes or not get_value(doc, "ActivateView", get_value(view, "GetName2")):
        return 0, 0
    try:
        marked = [hole for hole in holes if _mark(doc, view, hole, scale)]
    finally:
        get_value(doc, "ClearSelection2", True)
        get_value(doc, "ActivateView", "")
    groups = [(group, _pitch_circle(group, marked, scale)) for group in _groups(marked)]
    axes = tuple(axis_connections(marked, scale, groups=[group for group, circle in groups if circle is None]))
    circles = [circle for _, circle in groups if circle is not None]
    return len(marked), add_view_centerlines(doc, view, axes, circles=circles)


def _find_bore(annotation, nominal, data, matrix, bores, *, threaded=False):
    axis = axis_of(matrix[2:9:3])
    for raw in _optional(annotation, "GetAttachedEntities3") or ():
        edge = as_dispatch(raw)
        curve = _optional(edge, "GetCurve")
        params = _optional(curve, "CircleParams") if curve is not None else None
        if not params or len(params) < 7 or axis_of(params[3:6]) != axis:
            continue
        center = tuple(value for i, value in enumerate(params[:3]) if i != axis)
        matches = [b for b in bores if b.axis == axis and dist(b.center, center) <= TOLERANCE
                   and abs(b.radius-params[6]) <= TOLERANCE
                   and (threaded or abs(2*b.radius-nominal) <= TOLERANCE)
                   and b.low-TOLERANCE <= params[axis] <= b.high+TOLERANCE]
        if len(matches) == 1:
            return matches[0]
    return (_projected_bore(data, nominal, matrix, bores) or
            _linear_diameter_bore(annotation, nominal, data, matrix, bores, axis))


def _linear_diameter_bore(annotation, nominal, data, matrix, bores, axis):
    # SolidWorks can display a diameter between parallel extension lines.
    # In this style the midpoint of its arrows has moved off the circle.
    if _optional(data, "GetArrowHeadCount") != 2:
        return None
    tips = [_optional(data, "GetArrowHeadAtIndex2", i) for i in range(2)]
    if not all(tip and len(tip) >= 2 for tip in tips):
        return None
    scale = matrix[12]
    tolerance = max(4*TOLERANCE*scale, .0001)
    spans = [abs(tips[0][i]-tips[1][i]) for i in (0, 1)]
    along = 0 if spans[0] > spans[1] else 1
    if spans[1-along] > tolerance or abs(spans[along]-nominal*scale) > tolerance:
        return None
    strokes = dimension_lines(annotation)
    extensions = [point for a, b in strokes if abs(a[along]-b[along]) < tolerance
                  for point in (a, b)]
    matches = []
    for bore in bores:
        if bore.axis != axis or abs(2*bore.radius-nominal) > TOLERANCE:
            continue
        center = project(bore.point(bore.low), matrix)
        if all(any(abs(p[along]-tip[along]) <= tolerance and
                   abs(p[1-along]-center[1-along]) <= tolerance for p in extensions) and
               abs(abs(tip[along]-center[along])-bore.radius*scale) <= tolerance for tip in tips):
            matches.append(bore)
    return matches[0] if len(matches) == 1 else None


def _diameter_label(display, annotation, dimension, matrix, bores):
    value = _optional(dimension, "GetSystemValue2", "")
    data = _optional(annotation, "GetDisplayData")
    _, _, text = text_geometry(annotation, display_data=data)
    content = "".join(text).strip() if text and all(isinstance(part, str) for part in text) else ""
    match = fullmatch(r"(?:<MOD-DIAM>|[Ø⌀∅])\s*(\d+(?:[.,]\d+)?)", content)
    thread = fullmatch(r"[MМ]\d+(?:[.,]\d+)?", content)
    if not isinstance(value, (int, float)) or (not match and not thread):
        return None
    if match and not isclose(float(match[1].replace(',', '.'))/1000, value, abs_tol=TOLERANCE):
        return None
    if match and _optional(display, "Type2") != DIAMETER:
        return None
    bore = _find_bore(annotation, value, data, matrix, bores, threaded=bool(thread))
    if bore is not None or thread:
        return bore, content if thread else ""
    return None


def _place_quantity(layout, item, count):
    prefix = _optional(item.display, "GetText", 1) or ""
    before = "".join(text_geometry(item.annotation)[2])
    count_text = f"{count} отв. "
    get_value(item.display, "SetText", 1, count_text + prefix)
    applied = False
    try:
        after = "".join(text_geometry(item.annotation)[2])
        if "".join(after.split()) == "".join((count_text + before).split()):
            applied = layout.place(item.view, item.annotation, item.kind)
        return applied
    finally:
        if not applied:
            get_value(item.display, "SetText", 1, prefix)


def _covered_bores(group, bore):
    if bore is None:
        return frozenset(part for hole in group for part in hole.bores)
    profile = _stations((bore.radius, bore.low, bore.high))
    return frozenset(part for hole in group for part in hole.bores
                     if _stations((part.radius, part.low, part.high)) == profile)


def _centerline_strokes(holes, scale):
    plans = [(group, _pitch_circle(group, holes, scale)) for group in _groups(holes)]
    rows = [group for group, circle in plans if circle is None]
    for a, b in axis_connections(holes, scale, groups=rows):
        yield a[:2], b[:2]
    for _, circle in plans:
        if circle is not None:
            (x, y, _), radius = circle
            points = [(x+radius*cos(i*tau/64), y+radius*sin(i*tau/64)) for i in range(65)]
            yield from zip(points, points[1:])
    for hole in holes:
        x, y = hole.center
        span = hole.rim.radius*scale + .0015
        yield (x-span, y), (x+span, y)
        yield (x, y-span), (x, y+span)


def _is_stepped(group, bore):
    return any(abs(part.radius-bore.radius) > TOLERANCE for part in group[0].bores)


def _arrows_span_hole(annotation, center, radius):
    data = _optional(annotation, "GetDisplayData")
    if _optional(data, "GetArrowHeadCount") != 2:
        return False
    tips = [_optional(data, "GetArrowHeadAtIndex2", index) for index in range(2)]
    if not all(tip and len(tip) >= 2 for tip in tips):
        return False
    a, b = (tip[:2] for tip in tips)
    midpoint = tuple((x+y)/2 for x, y in zip(a, b))
    return dist(midpoint, center) <= .00015 and abs(dist(a, b)-2*radius) <= .00015


def _route_group_diameter(doc, layout, view, item, group, bore, matrix, holes, *, used_centers, thread=""):
    # All candidates belong to the very same verified hole family. Keep
    # the diameter associative by selecting a visible edge of that bore.
    members = _covered_bores(group, bore)
    scale = matrix[12]
    anchors = []
    for hole in group:
        if thread:
            edges = (hole.rim,) if hole.rim.fraction >= .80 else ()
        else:
            edges = [edge for edge in hole.edges if any(
                abs(edge.radius-member.radius) <= TOLERANCE
                and member.low-TOLERANCE <= edge.anchor.model[member.axis] <= member.high+TOLERANCE
                for member in hole.bores if member in members)]
        if edges:
            circle = max(edges, key=lambda edge: (edge.fraction, edge.anchor.sheet[2]))
            anchors.append((circle, hole.center, hole.rim.radius*scale))
    box, angle, _ = text_geometry(item.annotation)
    if not anchors or box is None:
        return None
    width, height = box[2]-box[0], box[3]-box[1]
    if angle is not None and abs(sin(angle)) > abs(cos(angle)):
        width, height = height, width
    count_text = f"{len(group)} отв. "
    label_size = width + len(count_text)*height*.55, height
    obstacles = _centerline_strokes(holes, scale)
    original = item.annotation if get_value(view, "GetName2") == get_value(item.view, "GetName2") else None
    for circle, position, center, radius, accept in layout.free_diameter_routes(
            view, original, anchors, label_size, obstacles, used_centers=used_centers,
            allow_outline_exit=bool(thread) or _is_stepped(group, bore), refine=bool(thread)):
        nominal = 2*(circle.radius if thread else bore.radius)
        annotation = _create(doc, view, Task((circle.anchor,), DIAMETER, nominal, 0), position, matrix)
        if annotation is None:
            continue
        display = get_value(annotation, "GetSpecificAnnotation")
        retained = False
        try:
            display.DisplayAsLinear = False
            get_value(display, "SetSecondArrow", False, bool(thread))
            if thread:
                # Keep the real opening as the attachment. Only the label is overridden.
                get_value(display, "SetText", 0, count_text + thread)
            else:
                prefix = _optional(display, "GetText", 1) or "<MOD-DIAM>"
                get_value(display, "SetText", 1, count_text + prefix)
            for inside in ((False, True) if thread else (False,)):
                if inside:
                    display.ArrowSide = 0
                get_value(doc, "GraphicsRedraw2")
                box, _, parts = text_geometry(annotation)
                text = "".join(parts).strip()
                if thread:
                    valid = text == count_text + thread and _arrows_span_hole(annotation, center, radius)
                else:
                    match = fullmatch(r"\d+\s+отв\.\s*(?:<MOD-DIAM>|[Ø⌀∅])\s*(\d+(?:[.,]\d+)?)", text)
                    valid = (match and text.startswith(count_text.strip()) and
                             isclose(float(match[1].replace(",", "."))/1000, nominal, abs_tol=TOLERANCE))
                if valid and accept(annotation, center, radius) and _hide(item.annotation):
                    layout.occupied.pop(layout._key(item.view, item.annotation), None)
                    layout.occupied[layout._key(view, annotation)] = box, dimension_lines(annotation)
                    used_centers.add(center)
                    retained = True
                    return annotation
        finally:
            if not retained:
                _discard(doc, view, display, annotation)
    return None


def _place_view_diameter(doc, layout, item, group, bore, matrix, holes, used_centers):
    center = next(hole.center for hole in group if bore in hole.bores)
    source = str(_optional(item.dimension, "FullName") or "")
    # Model-owned diameter labels can lose their added quantity after reopening.
    # For stepped holes prefer a drawing-owned leader, or reuse an uncrowded native one.
    prefer_original = center not in used_centers and (not _is_stepped(group, bore) or source.lower().endswith(".drawing"))
    if not prefer_original:
        callout = _route_group_diameter(doc, layout, item.view, item, group, bore, matrix, holes,
                                       used_centers=used_centers)
        if callout is not None:
            return callout
    if _place_quantity(layout, item, len(group)):
        used_centers.add(center)
        return item.annotation
    if prefer_original:
        return _route_group_diameter(doc, layout, item.view, item, group, bore, matrix, holes,
                                     used_centers=used_centers)
    return None


def _center_thread_note(doc, layout, view, note, hole, matrix, label):
    if hole.rim.fraction < .80:
        return note
    key = layout._key(view, note)
    box, previous = layout.occupied[key]
    edges = layout.edges.get(get_value(view, "GetName2"), ())
    other_lines = tuple(line for other, (_, lines) in layout.occupied.items() if other != key for line in lines)
    # In a crowded projection, use the space already found for the note.
    # The centered leader must introduce no additional crossings there.
    original_crossings = _line_crossings(previous, edges), _line_crossings(previous, other_lines)
    position = ((box[0]+box[2])/2, box[3]-.0008)
    rim = hole.rim
    annotation = _create(doc, view, Task((rim.anchor,), DIAMETER, 2*rim.radius, 0), position, matrix)
    if annotation is None:
        return note
    display = get_value(annotation, "GetSpecificAnnotation")
    occupied = layout.occupied.pop(key)
    kept = False
    try:
        display.DisplayAsLinear = False
        get_value(display, "SetSecondArrow", False, True)
        get_value(display, "SetText", 0, label)
        for inside in (False, True):
            if inside:
                display.ArrowSide = 0
            get_value(doc, "GraphicsRedraw2")
            if ("".join(text_geometry(annotation)[2]).strip() != label
                    or not layout.place(view, annotation, DIAMETER)
                    or not _arrows_span_hole(annotation, hole.center, rim.radius*matrix[12])):
                continue
            lines = dimension_lines(annotation)
            if (_line_crossings(lines, edges) <= original_crossings[0]
                    and _line_crossings(lines, other_lines) <= original_crossings[1]
                    and _hide(note)):
                kept = True
                return annotation
        return note
    finally:
        if not kept:
            layout.occupied.pop(layout._key(view, annotation), None)
            layout.occupied[key] = occupied
            _discard(doc, view, display, annotation)


def _place_thread_group(doc, layout, item, group, thread, matrix):
    # A model-owned thread dimension can reset its text when the drawing is
    # reopened. Create a drawing-owned leader attached to a verified rim.
    origin = _optional(item.annotation, "GetPosition")
    if not origin or not get_value(doc, "ActivateView", get_value(item.view, "GetName2")):
        return None
    hole = min(group, key=lambda hole: dist(hole.center, origin[:2]))
    label = f"{len(group)} отв. {thread}"
    key = layout._key(item.view, item.annotation)
    occupied = layout.occupied.pop(key, None)
    annotation = None
    kept = False
    try:
        get_value(doc, "ClearSelection2", True)
        if not _optional(item.view, "SelectEntity", hole.rim.anchor.entity, False):
            return None
        note = _optional(doc, "InsertNote", label)
        if note is None:
            return None
        annotation = _optional(note, "GetAnnotation")
        if (annotation is None or _optional(annotation, "GetAttachedEntityCount3") != 1
                or not _optional(annotation, "GetLeaderCount") or _optional(note, "GetText") != label
                or not _optional(annotation, "SetPosition2", *origin[:2], 0.)):
            return None
        annotation.Color = 0
        kept = layout.place(item.view, annotation, DIAMETER) and _hide(item.annotation)
        if not kept:
            return None
        try:
            return _center_thread_note(doc, layout, item.view, annotation, hole, matrix, label)
        except (AttributeError, com_error):
            return annotation
    finally:
        get_value(doc, "ClearSelection2", True)
        if not kept:
            if occupied is not None:
                layout.occupied[key] = occupied
            if annotation is not None:
                layout.occupied.pop(layout._key(item.view, annotation), None)
                if not (_optional(annotation, "Select3", False, VARIANT(VT_DISPATCH, None))
                        and _optional(get_value(doc, "Extension"), "DeleteSelection2", 0)):
                    _hide(annotation)
        get_value(doc, "ClearSelection2", True)
        get_value(doc, "ActivateView", "")


def _add_view_quantities(doc, view, holes, matrix, layout, used_centers):
    """Keep the native dimension and prefix its verified number of holes."""
    groups = _groups(holes)
    applicable = {(bore, hole.thread): group for group in groups for hole in group
                  if hole.thread or not hole.profile[2] for bore in hole.bores}
    thread_groups = {g[0].thread: g for g in groups if g[0].thread and
                     sum(hole.thread == g[0].thread for hole in holes) == len(g)}
    bores = tuple(dict.fromkeys(bore for hole in holes for bore in hole.bores))
    assigned = set()
    placed = []
    covered_tap_drills = set()
    hidden = 0
    name = get_value(view, "GetName2")
    for item in layout.items:
        if get_value(item.view, "GetName2") != name or item.dimension is None:
            continue
        label = _diameter_label(item.display, item.annotation, item.dimension, matrix, bores)
        if label is None:
            continue
        bore, thread = label
        thread_name = _thread_name(thread)
        group = applicable.get((bore, thread_name)) if bore is not None else thread_groups.get(thread_name)
        if group is None:
            continue
        key = (group[0].profile, thread_name or _stations((bore.radius, bore.low, bore.high)))
        if key in assigned:
            if bore is not None and _hide(item.annotation):
                layout.occupied.pop(layout._key(item.view, item.annotation), None)
                hidden += 1
            continue
        if thread:
            try:
                annotation = _route_group_diameter(doc, layout, view, item, group, None, matrix, holes,
                                                   used_centers=used_centers, thread=thread)
            except (AttributeError, com_error):
                annotation = None
            if annotation is None:
                annotation = _place_thread_group(doc, layout, item, group, thread, matrix)
        else:
            annotation = _place_view_diameter(doc, layout, item, group, bore, matrix, holes, used_centers)
        if annotation is not None:
            assigned.add(key)
            placed.append((_covered_bores(group, None if thread else bore), thread_name, annotation))
            if thread:
                covered_tap_drills.update(hole.tap_drill for hole in group if hole.tap_drill is not None)
    return placed, hidden, covered_tap_drills


def _hide_redundant_tap_drill_diameters(layout, hole_views, covered_tap_drills):
    views = {}
    for view, holes, matrix, flat in hole_views:
        if flat or get_value(view, "Type") == 2:
            continue
        bores = {bore for hole in holes for bore in hole.bores}
        if not covered_tap_drills.isdisjoint(bores):
            views[get_value(view, "GetName2")] = matrix, tuple(bores)
    hidden = 0
    for item in layout.items:
        projection = views.get(get_value(item.view, "GetName2"))
        if projection is None or item.kind != DIAMETER or _optional(item.annotation, "Visible") != VISIBLE:
            continue
        name = str(_optional(item.dimension, "FullName") or "")
        if not name.lower().endswith(".drawing") or _optional(item.annotation, "IsDangling"):
            continue
        matrix, bores = projection
        label = _diameter_label(item.display, item.annotation, item.dimension, matrix, bores)
        if label is not None and not label[1] and label[0] in covered_tap_drills and _hide(item.annotation):
            layout.occupied.pop(layout._key(item.view, item.annotation), None)
            hidden += 1
    return hidden


def add_hole_quantities(doc, frame, hole_views, cut_diameters):
    end_projections = [(view, holes, matrix) for view, holes, matrix, flat in hole_views
                       if not flat and get_value(view, "Type") != 2 and _groups(holes)]
    if not end_projections:
        return 0, 0
    layout = DimensionTextLayout(doc, frame)
    end_views = defaultdict(list)
    used_centers = defaultdict(set)
    placed = []
    covered_tap_drills = set()
    hidden = 0
    for view, holes, matrix in end_projections:
        current, removed, taps = _add_view_quantities(doc, view, holes, matrix, layout,
                                                      used_centers[get_value(view, "GetName2")])
        placed.extend(current)
        hidden += removed
        covered_tap_drills.update(taps)
        for group in _groups(holes):
            if group[0].profile[2]:  # cosmetic threads have their own annotation
                continue
            for bore in _covered_bores(group, None):
                end_views[bore].append((view, holes, matrix, group))
    by_name = {get_value(item.annotation, "GetName"): item for item in layout.items
               if get_value(item.view, "Type") == 2}
    for bore, annotation in cut_diameters:
        candidates = end_views.get(bore)
        item = by_name.get(get_value(annotation, "GetName")) if candidates else None
        if item is None or not _plain_diameter(_optional(annotation, "GetDisplayData"), 2*bore.radius):
            continue
        if any(not thread and bore in footprint for footprint, thread, _ in placed):
            if _hide(annotation):
                layout.occupied.pop(layout._key(item.view, annotation), None)
                hidden += 1
            continue
        # A cut diameter can supply the nominal size. Its quantity belongs
        # with a visible circular opening of the same bore in an end view.
        for view, holes, matrix, group in sorted(candidates, key=lambda row: -len(row[3])):
            callout = _route_group_diameter(doc, layout, view, item, group, bore, matrix, holes,
                                           used_centers=used_centers[get_value(view, "GetName2")])
            if callout is not None:
                placed.append((_covered_bores(group, bore), "", callout))
                break
    # The reverse projection can show a complete group while the front
    # exposes only part of it. Suppress the lesser label only when its
    # physical holes are all covered by another successful callout.
    retained = []
    redundant = 0
    for footprint, thread, annotation in sorted(placed, key=lambda row: -len(row[0])):
        if footprint and any(thread == other_thread and footprint <= other for other, other_thread in retained):
            if _hide(annotation):
                redundant += 1
                continue
        retained.append((footprint, thread))
    if covered_tap_drills:
        hidden += _hide_redundant_tap_drill_diameters(layout, hole_views, covered_tap_drills)
    return len(placed)-redundant, hidden+redundant
