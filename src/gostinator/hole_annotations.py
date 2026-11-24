"""Center marks and shared diameter callouts for visible holes."""
from collections import defaultdict
from dataclasses import dataclass
from math import atan2, cos, dist, isclose, sin, tau
from re import fullmatch

from pythoncom import VT_DISPATCH
from win32com.client import VARIANT

from .annotation_layout import DimensionTextLayout, dimension_lines, text_geometry
from .drawing_dimensions import (
    Circle, DIAMETER, Task, _create, _discard, _hide, _optional, _plain_diameter, _projected_bore,
)
from .drawing_views import add_view_centerlines, project
from .model_geometry import POSITION_TOLERANCE as TOLERANCE, axis_of
from .solidworks import as_dispatch, get_value


@dataclass(frozen=True)
class Hole:
    center: tuple[float, float]
    rim: Circle
    bores: tuple
    profile: tuple
    thread: str = ""
    edges: tuple[Circle, ...] = ()


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
            spec = (_thread_name(size) if isinstance(size, str) else "",
                    round(diameter/TOLERANCE) if isinstance(diameter, (float, int)) else None,
                    round(depth/TOLERANCE) if isinstance(depth, (float, int)) else None, fit)
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
        thread_profile = sources[0] if len(set(sources)) == 1 else ("", None, None, None)
        result.append(Hole(center, rim, tuple(bores), (profile, visible_rings, tuple(sorted(set(sources), key=str))),
                           thread_profile[0], tuple(matched)))
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


def _relocate_group_diameter(doc, layout, item, group, bore, matrix, holes):
    # All candidates belong to the very same verified hole family. Keep
    # the diameter associative by selecting a visible edge of that bore.
    members = _covered_bores(group, bore)
    scale = matrix[12]
    anchors = []
    for hole in group:
        edges = [edge for edge in hole.edges if any(
            abs(edge.radius-member.radius) <= TOLERANCE
            and member.low-TOLERANCE <= edge.anchor.model[member.axis] <= member.high+TOLERANCE
            for member in hole.bores if member in members)]
        if edges:
            circle = max(edges, key=lambda edge: (edge.fraction, edge.anchor.sheet[2]))
            anchors.append((circle.anchor, hole.center, hole.rim.radius*scale))
    box, _, _ = text_geometry(item.annotation)
    if not anchors or box is None:
        return None
    width, height = box[2]-box[0], box[3]-box[1]
    count_text = f"{len(group)} отв. "
    label_size = width + len(count_text)*height*.55, height
    obstacles = _centerline_strokes(holes, scale)
    nominal = 2*bore.radius
    for anchor, position, center, radius, accept in layout.free_diameter_routes(
            item.view, item.annotation, anchors, label_size, obstacles):
        annotation = _create(doc, item.view, Task((anchor,), DIAMETER, nominal, 0), position, matrix)
        if annotation is None:
            continue
        display = get_value(annotation, "GetSpecificAnnotation")
        retained = False
        try:
            display.DisplayAsLinear = False
            get_value(display, "SetSecondArrow", False, False)
            prefix = _optional(display, "GetText", 1) or "<MOD-DIAM>"
            get_value(display, "SetText", 1, count_text + prefix)
            get_value(doc, "GraphicsRedraw2")
            box, _, text = text_geometry(annotation)
            text = "".join(text).strip()
            match = fullmatch(r"\d+\s+отв\.\s*(?:<MOD-DIAM>|[Ø⌀∅])\s*(\d+(?:[.,]\d+)?)", text)
            if (match and text.startswith(count_text.strip())
                    and isclose(float(match[1].replace(",", "."))/1000, nominal, abs_tol=TOLERANCE)
                    and accept(annotation, center, radius) and _hide(item.annotation)):
                layout.occupied.pop(layout._key(item.view, item.annotation), None)
                layout.occupied[layout._key(item.view, annotation)] = box, dimension_lines(annotation)
                retained = True
                return annotation
        finally:
            if not retained:
                _discard(doc, item.view, display, annotation)
    return None


def _place_thread_group(doc, layout, item, group, thread):
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
        return annotation if kept else None
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


def _add_view_quantities(doc, view, holes, matrix, layout):
    """Keep the native dimension and prefix its verified number of holes."""
    groups = _groups(holes)
    applicable = {(bore, hole.thread): group for group in groups for hole in group
                  if hole.thread or not hole.profile[2] for bore in hole.bores}
    thread_groups = {g[0].thread: g for g in groups if g[0].thread and
                     sum(hole.thread == g[0].thread for hole in holes) == len(g)}
    bores = tuple(dict.fromkeys(bore for hole in holes for bore in hole.bores))
    assigned = set()
    placed = []
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
            annotation = _place_thread_group(doc, layout, item, group, thread)
        elif _place_quantity(layout, item, len(group)):
            annotation = item.annotation
        else:
            annotation = _relocate_group_diameter(doc, layout, item, group, bore, matrix, holes)
        if annotation is not None:
            assigned.add(key)
            placed.append((_covered_bores(group, None if thread else bore), thread_name, annotation))
    return placed, hidden


def add_hole_quantities(doc, frame, hole_views, cut_diameters):
    if not any(not flat and _groups(holes) for _, holes, _, flat in hole_views):
        return 0, 0
    layout = DimensionTextLayout(doc, frame)
    memberships = defaultdict(set)
    placed = []
    hidden = 0
    for view, holes, matrix, flat in hole_views:
        if flat:
            continue
        current, removed = _add_view_quantities(doc, view, holes, matrix, layout)
        placed.extend(current)
        hidden += removed
        for group in _groups(holes):
            if group[0].profile[2]:  # cosmetic threads have their own annotation
                continue
            for hole in group:
                for bore in hole.bores:
                    memberships[bore].add((len(group), _covered_bores(group, bore)))
    by_name = {get_value(item.annotation, "GetName"): item for item in layout.items
               if get_value(item.view, "Type") == 2}
    for bore, annotation in cut_diameters:
        groups = memberships.get(bore)
        item = by_name.get(get_value(annotation, "GetName")) if groups else None
        if item is not None and len(groups) == 1 and _plain_diameter(
                _optional(annotation, "GetDisplayData"), 2*bore.radius):
            count, footprint = next(iter(groups))
            if _place_quantity(layout, item, count):
                placed.append((footprint, "", item.annotation))
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
    return len(placed)-redundant, hidden+redundant
