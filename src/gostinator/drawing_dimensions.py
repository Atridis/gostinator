"""Select and dimension visible model geometry on the native drawing."""
from collections import Counter, defaultdict
from dataclasses import dataclass, replace
from math import acos, dist, isclose, isfinite, sqrt, tau
from re import fullmatch

from pythoncom import VT_DISPATCH, com_error
from win32com.client import VARIANT

from .annotation_layout import geometry_outline, view_strokes, visible_dimensions
from .drawing_views import drawing_views, project, transform_of
from .model_geometry import (
    Bore, POSITION_TOLERANCE as TOLERANCE, axis_of, choose_section, perpendicular_flange_profile, primary_axis,
)
from .solidworks import as_dispatch, get_value


LINEAR = 2
ANGLE = 3
RADIUS = 5
DIAMETER = 6
VISIBLE, HIDDEN = 1, 3
SW_ARC_END_FARTHEST = 3
SW_DIMENSION_TEXT_PREFIX = 1
SW_END_BLIND = 0
SW_END_MIDPLANE = 6


@dataclass(frozen=True)
class Anchor:
    entity: object
    model: tuple[float, float, float]
    sheet: tuple[float, float, float]
    silhouette: bool = False


@dataclass(frozen=True)
class Circle:
    anchor: Anchor
    axis: int
    radius: float
    fraction: float
    ends: tuple = ()


@dataclass(frozen=True)
class Line:
    entity: object
    ends: tuple[Anchor, Anchor]
    silhouette: bool = False


@dataclass(frozen=True)
class Task:
    targets: tuple[Anchor, ...]
    kind: int
    nominal: float
    direction: int
    projected: bool = False
    overall: bool = False
    bore: Bore | None = None
    profile: bool = False
    arc_ends: int = 0
    position: tuple[float, float] | None = None


def _optional(obj, name, *args):
    try:
        return get_value(obj, name, *args)
    except (AttributeError, com_error):
        return None


def _visible(view, kind):
    values = _optional(view, "GetVisibleEntities2", None, kind)
    if values is None:
        values = _optional(view, "GetVisibleEntities2", VARIANT(VT_DISPATCH, None), kind)
    return tuple(map(as_dispatch, values or ()))


def _section_silhouettes(view, matrix):
    for edge in _visible(view, 4):
        curve = _optional(edge, "GetCurve")
        if curve is None or not _optional(curve, "IsLine"):
            continue
        points = [_optional(edge, method) for method in ("GetStartPoint", "GetEndPoint")]
        coords = [_optional(point, "ArrayData") if point is not None else None for point in points]
        if all(point and len(point) >= 3 for point in coords):
            yield Line(edge, tuple(Anchor(edge, tuple(p[:3]), project(p, matrix), True)
                                   for p in coords), True)


def _point_key(point):
    return tuple(round(value / TOLERANCE) for value in point[:3])


def _anchors(view, matrix):
    vertices = {}
    circles = []
    lines = []
    for vertex in _visible(view, 2):
        point = _optional(vertex, "GetPoint")
        if point and len(point) >= 3:
            vertices[_point_key(point)] = Anchor(vertex, tuple(point[:3]), project(point, matrix))
    for edge in _visible(view, 1):
        curve = _optional(edge, "GetCurve")
        if curve is None:
            continue
        if _optional(curve, "IsCircle"):
            data = _optional(curve, "CircleParams")
            if data and len(data) >= 7 and data[6] > TOLERANCE:
                axis = axis_of(data[3:6])
                if axis is not None:
                    params = _optional(edge, "GetCurveParams3")
                    lo = _optional(params, "UMinValue") if params is not None else None
                    hi = _optional(params, "UMaxValue") if params is not None else None
                    fraction = abs(hi - lo) / tau if lo is not None and hi is not None else 1.0
                    ends = tuple(_optional(params, name) for name in ("StartPoint", "EndPoint")
                                 ) if params is not None and .12 <= fraction <= .54 else ()
                    circles.append(Circle(Anchor(edge, tuple(data[:3]), project(data, matrix)),
                                          axis, data[6], fraction, ends))
        elif _optional(curve, "IsLine"):
            ends = []
            for getter in ("GetStartVertex", "GetEndVertex"):
                vertex = _optional(edge, getter)
                point = _optional(vertex, "GetPoint") if vertex is not None else None
                if point and len(point) >= 3:
                    ends.append(vertices.setdefault(
                        _point_key(point), Anchor(vertex, tuple(point[:3]), project(point, matrix))))
            if len(ends) == 2:
                lines.append(Line(edge, tuple(ends)))
    return list(vertices.values()), circles, lines


def _key(kind, value):
    return (LINEAR if kind in (11, 12) else kind, round(value / TOLERANCE))


def _existing(view):
    used = set()
    for display in _optional(view, "GetDisplayDimensions") or ():
        dimension = _optional(display, "GetDimension2", 0)
        annotation = _optional(display, "GetAnnotation")
        if dimension is None or annotation is None or _optional(annotation, "Visible") != VISIBLE:
            continue
        value = _optional(dimension, "GetSystemValue2", "")
        kind = _optional(display, "Type2")
        if kind is not None and isinstance(value, (int, float)):
            used.add(_key(kind, value))
    return used


def _geometry_key(task, matrix, normal, view_scope=None):
    kind, value = _key(task.kind, task.nominal)
    world_axis = axis_of(matrix[task.direction:9:3]) if kind == LINEAR else None
    if task.overall and world_axis is not None and view_scope is None:
        extent = tuple(sorted(round(a.model[world_axis] / TOLERANCE) for a in task.targets))
        return kind, value, world_axis, extent
    # Opposite ordinary projections can show the same feature. Drop depth
    # when comparing their projected attachment positions, but keep the
    # measuring axis and the other coordinates of both attachments.
    anchors = tuple(sorted(tuple(round(x / TOLERANCE) for i, x in enumerate(a.model)
                                 if i != normal) for a in task.targets))
    return view_scope, kind, value, normal, world_axis, anchors


def _folded_profile_shift(vertices, model, normal, position):
    if not vertices:
        return None
    limits = [(min(values), max(values)) for values in zip(*(v.model for v in vertices))]
    tolerance = 4 * TOLERANCE
    if all(abs(high - low - model.spans[i]) <= tolerance
           for i, (low, high) in enumerate(limits)):
        return tuple(low - model.bounds[i] for i, (low, _) in enumerate(limits))
    # SolidWorks can return the drawing body's vertices with the view's
    # sheet-position translation. Only use that origin if the two visible
    # stock extents match and every sampled vertex lies inside the body.
    shifts = [(0., 0., 0.)]
    if position and len(position) >= 2 and dist(position[:2], (0., 0.)) > tolerance:
        shifts.append((*position[:2], 0.))
    matches = []
    for shift in shifts:
        gaps = [(low - shift[i] - model.bounds[i], model.bounds[i+3] - (high - shift[i]))
                for i, (low, high) in enumerate(limits)]
        if all(min(pair) >= -tolerance and (i == normal or max(pair) <= tolerance)
               for i, pair in enumerate(gaps)):
            matches.append(shift)
    return matches[0] if len(matches) == 1 else None


def _registered_vertices(vertices, lines, matrix, shift):
    def register(anchor):
        point = tuple(value - delta for value, delta in zip(anchor.model, shift))
        return replace(anchor, model=point, sheet=project(point, matrix))

    return ([register(vertex) for vertex in vertices],
            [replace(line, ends=tuple(map(register, line.ends))) for line in lines])


def _arc_vertex_shift(vertices, circles, position):
    # Drawing vertices can include the view position while the corresponding
    # curve parameters are returned in model coordinates.
    if not position or dist(position[:2], (0., 0.)) <= TOLERANCE:
        return None
    ends = {_point_key(point) for circle in circles for point in circle.ends if point}
    if len(ends) < 2:
        return None
    shifts = ((0., 0., 0.), (*position[:2], 0.))
    scores = [len(ends & {_point_key(tuple(a-b for a, b in zip(v.model, shift))) for v in vertices})
              for shift in shifts]
    if scores[1] >= 2 and scores[0] == 0:
        return shifts[1]
    return shifts[0] if scores[0] >= 2 and scores[1] == 0 else None


def _physical_profile_key(task, matrix, shift):
    axis = axis_of(matrix[task.direction:9:3])
    if axis is None:
        return None
    points = (_point_key(tuple(value - delta for value, delta in zip(anchor.model, shift)))
              for anchor in task.targets)
    return axis, tuple(sorted(points))


def _profile_interval(profile):
    axis, ends = profile
    return axis, tuple(sorted(point[axis] for point in ends))


def _cut_depth(view, dimension, name):
    """Return the value only for a verified depth of an extruded cut."""
    if not isinstance(name, str) or not name.casefold().endswith(".part"):
        return None
    owner = _optional(dimension, "GetFeatureOwner")
    if owner is None:
        components = name.split("@", 2)
        source = _optional(view, "ReferencedDocument")
        if len(components) == 3 and source is not None:
            owner = _optional(source, "FeatureByName", components[1])
    if owner is None:
        return None
    kind = _optional(owner, "GetTypeName2")
    if kind == "ICE":
        kind = _optional(owner, "GetTypeName")
    if kind not in ("Cut", "CutThin"):
        return None
    definition = _optional(owner, "GetDefinition")
    value = _optional(dimension, "GetSystemValue2", "")
    if definition is None or not isinstance(value, (int, float)) or not isfinite(value) or value <= 0.:
        return None
    directions = (True, False) if _optional(definition, "BothDirections") else (True,)
    for forward in directions:
        if _optional(definition, "GetEndCondition", forward) not in (SW_END_BLIND, SW_END_MIDPLANE):
            continue
        depth = _optional(definition, "GetDepth", forward)
        if isinstance(depth, (int, float)) and isclose(value, depth, abs_tol=TOLERANCE, rel_tol=0.):
            return value
    return None


def _depth_extends_past_material(tips, strokes, nominal, scale):
    if len(tips) != 2 or not strokes or not all(
            tip and len(tip) >= 2 and all(map(isfinite, tip[:2])) for tip in tips):
        return False
    start, end = (tip[:2] for tip in tips)
    measured = dist(start, end)
    clearance = max(.0001, 4*TOLERANCE*scale)
    if measured <= clearance or not isclose(measured, nominal*scale, abs_tol=clearance, rel_tol=0.):
        return False
    axis = tuple((b-a)/measured for a, b in zip(start, end))

    def station(point):
        return sum(x*y for x, y in zip(point, axis))

    stations = [station(point) for stroke in strokes for point in stroke]
    low, high = min(stations)-clearance, max(stations)+clearance
    # Require one real material-side datum and one end beyond the displayed
    # part envelope. An incomplete projection cannot establish the depth.
    return (low <= station(start) <= high) != (low <= station(end) <= high)


def prune_model_dimensions(doc):
    """Hide duplicates, dangling marks and cut lengths extending past stock."""
    names = set()
    removed = 0
    zoomed = False
    for view in drawing_views(doc):
        strokes = None
        for display in _optional(view, "GetDisplayDimensions") or ():
            annotation = _optional(display, "GetAnnotation")
            if annotation is None or _optional(annotation, "Visible") != VISIBLE:
                continue
            dimension = _optional(display, "GetDimension2", 0)
            name = _optional(dimension, "FullName") if dimension is not None else None
            dangling = _optional(annotation, "IsDangling")
            if dangling or (name and name in names):
                removed += _hide(annotation)
                continue
            if dimension is not None and _optional(display, "Type2") in (LINEAR, 11, 12):
                nominal = _cut_depth(view, dimension, name)
                if nominal is not None:
                    if not zoomed:
                        get_value(doc, "ViewZoomtofit2")
                        zoomed = True
                    data = _optional(annotation, "GetDisplayData")
                    if data is not None and _optional(data, "GetArrowHeadCount") == 2:
                        tips = [_optional(data, "GetArrowHeadAtIndex2", i) for i in range(2)]
                        if strokes is None:
                            box = get_value(view, "GetOutline")[:4]
                            strokes = view_strokes(view, box)
                        scale = _optional(view, "ScaleDecimal")
                        if scale and _depth_extends_past_material(tips, strokes, nominal, scale) and _hide(annotation):
                            removed += 1
                            continue
            if name:
                names.add(name)
    return removed


def _hide(annotation):
    try:
        annotation.Visible = HIDDEN
        return _optional(annotation, "Visible") == HIDDEN
    except (AttributeError, com_error):
        return False


def _section_bore_keys(section):
    coaxial = defaultdict(list)
    for bore in section.bores:
        coaxial[_point_key(bore.point(0.))].append(bore)
    keys = {}
    for group in coaxial.values():
        segments = {bore: _point_key((bore.radius, bore.low, bore.high)) for bore in group}
        profile = tuple(sorted(segments.values()))
        stepped = len({radius for radius, _, _ in profile}) > 1
        for bore, segment in segments.items():
            # Repeated stepped holes in the same cut can share a callout
            # only when the complete cylindrical profiles agree.
            keys[bore] = ((profile, segment) if stepped else bore, stepped)
    return keys


def _plain_diameter(data, nominal):
    if data is None:
        return False
    parts = [_optional(data, "GetTextAtIndex", i)
             for i in range(_optional(data, "GetTextCount") or 0)]
    if not all(isinstance(part, str) for part in parts):
        return False
    match = fullmatch(r"(?:<MOD-DIAM>|[Ø⌀∅])\s*(\d+(?:[.,]\d+)?)",
                      "".join(parts).strip())
    return bool(match and isclose(float(match[1].replace(",", ".")) / 1000, nominal,
                                  abs_tol=TOLERANCE))


def _projected_bore(data, nominal, matrix, bores):
    if _optional(data, "GetArrowHeadCount") != 2:
        return None
    tips = [_optional(data, "GetArrowHeadAtIndex2", i) for i in range(2)]
    if not all(tip and len(tip) >= 2 for tip in tips):
        return None
    scale = matrix[12]
    if not isclose(dist(tips[0][:2], tips[1][:2]), nominal * scale, abs_tol=TOLERANCE * scale):
        return None
    center = tuple((a + b) / 2 for a, b in zip(tips[0][:2], tips[1][:2]))
    axis = axis_of(matrix[2:9:3])
    matches = [bore for bore in bores
               if bore.axis == axis and isclose(2 * bore.radius, nominal, abs_tol=TOLERANCE)
               and dist(project(bore.point(bore.low), matrix)[:2], center) <= TOLERANCE * scale]
    return matches[0] if len(matches) == 1 else None


def prune_duplicate_bore_diameters(doc, model, cut_diameters):
    """Keep one plain diameter across a section and its ordinary projections."""
    if not cut_diameters:
        return 0
    section = choose_section(model, primary_axis(model))
    if section is None:
        return 0
    keys = _section_bore_keys(section)
    visible_cut = {}
    for bore, annotation in cut_diameters:
        if (_optional(annotation, "Visible") == VISIBLE and not _optional(annotation, "IsDangling")
                and _plain_diameter(_optional(annotation, "GetDisplayData"), 2 * bore.radius)):
            visible_cut[keys[bore]] = annotation
    if not visible_cut:
        return 0
    removed = 0
    for item in visible_dimensions(doc):
        if (item.kind != DIAMETER or item.dimension is None or _optional(item.view, "Type") == 2
                or _optional(item.view, "IsFlatPatternView") or _optional(item.annotation, "IsDangling")):
            continue
        data = _optional(item.annotation, "GetDisplayData")
        value = _optional(item.dimension, "GetSystemValue2", "")
        if not isinstance(value, (int, float)) or not _plain_diameter(data, value):
            continue
        bore = _projected_bore(data, value, transform_of(item.view), model.bores)
        key = keys.get(bore)
        cut = visible_cut.get(key)
        if cut is not None and _optional(cut, "Visible") == VISIBLE:
            removed += _hide(item.annotation if key[1] else cut)
    return removed


def _pair_on_axis(anchors, at_low, at_high, direction, tolerance):
    first = [a for a in anchors if abs(a.sheet[direction] - at_low) <= tolerance]
    second = [a for a in anchors if abs(a.sheet[direction] - at_high) <= tolerance]
    if not first or not second:
        return None
    along = 1 - direction
    a, b = min(((a, b) for a in first for b in second),
               key=lambda pair: (abs(pair[0].sheet[along] - pair[1].sheet[along]) > tolerance,
                                 abs(pair[0].sheet[2] - pair[1].sheet[2]),
                                 abs(pair[0].sheet[along] - pair[1].sheet[along])))
    if abs(a.sheet[along] - b.sheet[along]) <= tolerance and abs(a.sheet[2] - b.sheet[2]) <= tolerance:
        return a, b
    return None


def _overall_tasks(vertices, matrix, *, minimum=.003):
    if not vertices:
        return
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE, 1e-6)
    for direction in (0, 1):
        low = min(anchor.sheet[direction] for anchor in vertices)
        high = max(anchor.sheet[direction] for anchor in vertices)
        if (high - low) / scale < minimum:
            continue
        pair = _pair_on_axis(vertices, low, high, direction, tolerance)
        if pair:
            yield Task(pair, LINEAR, (high - low) / scale, direction, overall=True)
            continue
        # An irregular stock profile can reach its two true limits at
        # staggered corners. Dimension the axis projection, not the diagonal.
        a = min(vertices, key=lambda anchor: anchor.sheet[direction])
        b = max(vertices, key=lambda anchor: anchor.sheet[direction])
        yield Task((a, b), LINEAR, (high - low) / scale, direction, projected=True, overall=True)


def _blank_extent_tasks(vertices, matrix, contour, circles=(), lines=()):
    if contour is None:
        return
    done = set()
    for task in _overall_tasks(vertices, matrix, minimum=TOLERANCE):
        direction = task.direction
        span = (contour[direction+2] - contour[direction]) / matrix[12]
        # SolidWorks' drawing-view outline includes a margin. Compare the
        # selected endpoints with the actual displayed blank contour instead.
        if isclose(task.nominal, span, abs_tol=2*TOLERANCE, rel_tol=0.):
            done.add(direction)
            rails = _parallel_blank_sides(lines, contour, direction, matrix[12])
            yield replace(task, targets=rails or task.targets, projected=True)
    normal = axis_of(matrix[2:9:3])
    arcs = [circle for circle in circles if circle.axis == normal and .20 <= circle.fraction <= .55]
    for direction in (0, 1):
        if direction in done or not arcs:
            continue
        scale = matrix[12]
        point_ends = [(a.sheet[direction], a, 0) for a in vertices]
        near = point_ends + [(c.anchor.sheet[direction]-c.radius*scale, c.anchor, 1) for c in arcs]
        far = point_ends + [(c.anchor.sheet[direction]+c.radius*scale, c.anchor, 1) for c in arcs]
        low, a, first_arc = min(near, key=lambda item: item[0])
        high, b, second_arc = max(far, key=lambda item: item[0])
        measured = (high-low) / scale
        span = (contour[direction+2] - contour[direction]) / scale
        if a is not b and isclose(measured, span, abs_tol=2*TOLERANCE, rel_tol=0.):
            yield Task((a, b), LINEAR, measured, direction, projected=True, overall=True,
                       arc_ends=first_arc+second_arc)


def _parallel_blank_sides(lines, contour, direction, scale):
    # Opposite tangency points on a round-ended blank can trigger a
    # diameter display. Prefer the two parallel sides when they exist.
    tolerance = max(TOLERANCE*scale, 1e-6)
    along = 1-direction
    rails = []
    for side in (contour[direction], contour[direction+2]):
        candidates = [line for line in lines if all(abs(a.sheet[direction]-side) < tolerance for a in line.ends)
                      and abs(line.ends[0].sheet[along]-line.ends[1].sheet[along]) > .004*scale]
        if not candidates:
            return ()
        rails.append(candidates)
    pairs = [(a, b) for a in rails[0] for b in rails[1]
             if all(abs(first.sheet[2]-second.sheet[2]) <= tolerance
                    for first in a.ends for second in b.ends)]
    if not pairs:
        return ()
    selected = max(pairs, key=lambda pair: sum(dist(*(p.sheet for p in line.ends)) for line in pair))
    return tuple(Anchor(line.entity, line.ends[0].model,
                        tuple((a+b)/2 for a, b in zip(*(p.sheet for p in line.ends)))) for line in selected)


def _hole_tasks(circles, lines, model, view_axis, matrix, *, cap=None, vertices=()):
    matched = defaultdict(dict)
    for circle in circles:
        if circle.axis != view_axis:
            continue
        point = circle.anchor.model
        center = tuple(value for i, value in enumerate(point) if i != view_axis)
        bore = next((b for b in model.bores if b.axis == circle.axis
                     and dist(center, b.center) <= TOLERANCE and abs(b.radius-circle.radius) <= TOLERANCE
                     and b.low - .002 <= point[view_axis] <= b.high + .002), None)
        if bore is None:
            continue
        key = round(bore.radius / TOLERANCE)
        matched[key].setdefault(tuple(round(i / TOLERANCE) for i in center), circle)

    for group in matched.values():
        members = list(group.values())
        one = members[0]
        yield Task((one.anchor,), DIAMETER, 2 * one.radius, 0)
        # Dimension adjacent equal-size holes once per spacing. Only use
        # aligned centres for runs of holes, and two projected offsets for
        # a lone pair with different X and Y positions.
        for direction in (0, 1):
            transverse = 1 - direction
            by_row = defaultdict(list)
            tolerance = max(TOLERANCE * matrix[12], 1e-6)
            for circle in members:
                by_row[round(circle.anchor.sheet[transverse] / tolerance)].append(circle.anchor)
            for row in by_row.values():
                row.sort(key=lambda item: item.sheet[direction])
                for a, b in zip(row, row[1:]):
                    nominal = (b.sheet[direction] - a.sheet[direction]) / matrix[12]
                    if nominal >= .003:
                        yield Task((a, b), LINEAR, nominal, direction, projected=True)
            if len(members) == 2:
                a, b = (circle.anchor for circle in members)
                width = abs(a.sheet[direction] - b.sheet[direction]) / matrix[12]
                if abs(a.sheet[transverse] - b.sheet[transverse]) > tolerance and width >= .003:
                    yield Task((a, b), LINEAR, width, direction, projected=True)
            if cap is not None:
                datum = _rounded_flange_datum(members, vertices, cap, direction, matrix)
            else:
                datum = _datum_offset(members, lines, direction, matrix)
            if datum:
                yield datum


def _datum_offset(members, lines, direction, matrix):
    along = 1 - direction
    tolerance = max(TOLERANCE * matrix[12], 1e-6)
    real_extent = max(b.anchor.sheet[along] for b in members) - min(
        b.anchor.sheet[along] for b in members)
    choices = []
    envelope = tuple(operation(a.sheet[direction] for line in lines for a in line.ends)
                     for operation in (min, max)) if lines else ()
    for line in lines:
        a, b = line.ends
        if (abs(a.sheet[direction] - b.sheet[direction]) > tolerance
                or abs(a.sheet[along] - b.sheet[along]) < max(real_extent, .006 * matrix[12])
                or abs(a.sheet[2] - b.sheet[2]) > tolerance
                or not any(abs(a.sheet[direction] - outside) <= tolerance for outside in envelope)):
            continue
        low, high = sorted((a.sheet[along], b.sheet[along]))
        for circle in members:
            center = circle.anchor
            offset = abs(a.sheet[direction] - center.sheet[direction]) / matrix[12]
            if low - tolerance <= center.sheet[along] <= high + tolerance and offset > circle.radius + .001:
                midpoint = tuple((x + y) / 2 for x, y in zip(a.sheet, b.sheet))
                datum = Anchor(line.entity, a.model, midpoint)
                choices.append((offset, Task((center, datum), LINEAR, offset, direction, projected=True)))
    return min(choices, key=lambda pair: pair[0])[1] if choices else None


def _rounded_end(circles, normal, contour, scale):
    # A rounded flange end can be split into two quarter arcs. Its centre is
    # one radius in from just one end of the visible stock envelope.
    if contour is None:
        return None
    groups = defaultdict(list)
    for circle in circles:
        if circle.axis == normal and .20 <= circle.fraction <= .54:
            groups[(_point_key(circle.anchor.model), round(circle.radius/TOLERANCE))].append(circle)
    for group in groups.values():
        circle = group[0]
        radius = circle.radius * scale
        x, y = circle.anchor.sheet[:2]
        for direction, center in enumerate((x, y)):
            along = 1-direction
            if not isclose(contour[along+2]-contour[along], 2*radius, abs_tol=4*TOLERANCE):
                continue
            if not isclose((x, y)[along], (contour[along]+contour[along+2])/2, abs_tol=4*TOLERANCE):
                continue
            for side, sign in ((0, -1), (2, 1)):
                if isclose(center+sign*radius, contour[direction+side], abs_tol=4*TOLERANCE):
                    # Validate that the half-round actually reaches this end.
                    ends = [point for arc in group for point in arc.ends if point]
                    if len({_point_key(point) for point in ends}) >= 3 or any(
                            .44 <= arc.fraction <= .54 for arc in group):
                        return circle, direction, sign
    return None


def _rounded_flange_datum(members, vertices, cap, direction, matrix):
    _, along, sign = cap
    if direction != along or not vertices:
        return None
    center = min((c.anchor for c in members), key=lambda a: sign*a.sheet[direction])
    limit = min(sign*v.sheet[direction] for v in vertices)
    edge = min((v for v in vertices if abs(sign*v.sheet[direction]-limit) <= TOLERANCE),
               key=lambda v: (abs(v.sheet[2]-center.sheet[2]), abs(v.sheet[1-direction]-center.sheet[1])))
    nominal = abs(edge.sheet[direction]-center.sheet[direction])/matrix[12]
    if nominal >= .003:
        return Task((center, edge), LINEAR, nominal, direction, projected=True)
    return None


def _profile_tasks(lines, matrix):
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE, 1e-6)
    if not lines:
        return
    # A front-facing profile consists of real straight boundaries in the
    # same plane. Short trim and backing edges can misstate the profile.
    near = max(min(a.sheet[2] for a in line.ends) for line in lines)
    candidates = defaultdict(list)
    for line in lines:
        first, second = line.ends
        if min(first.sheet[2], second.sheet[2]) < near - tolerance:
            continue
        delta = [abs(a-b) for a, b in zip(first.sheet[:2], second.sheet[:2])]
        direction = 0 if delta[0] >= delta[1] else 1
        nominal = delta[direction] / scale
        if delta[1-direction] <= tolerance and nominal >= .004:
            candidates[direction].append(Task(line.ends, LINEAR, nominal, direction, profile=True))
    for tasks in candidates.values():
        # Allow one task per physical length family. Prefer longer, more
        # useful profile spans before short steps compete for space.
        yield from sorted(tasks, key=lambda t: t.nominal, reverse=True)


def _profile_radii(circles, lines, vertices, normal, matrix):
    if not vertices:
        return
    bounds = [(min(v.sheet[i] for v in vertices), max(v.sheet[i] for v in vertices)) for i in (0, 1)]
    for circle in circles:
        if circle.axis != normal or not .12 <= circle.fraction <= .54:
            continue
        radius = circle.radius * matrix[12]
        nearest = [min(abs(circle.anchor.sheet[i] - side) for side in bounds[i]) for i in (0, 1)]
        exterior_corner = .19 <= circle.fraction <= .32 and all(
            abs(gap - radius) <= max(.001 * matrix[12], radius * .25) for gap in nearest)
        if exterior_corner or _tangent_ends(circle, lines):
            yield Task((circle.anchor,), RADIUS, circle.radius, 0)


def _tangent_ends(circle, lines):
    return bool(_tangent_rails(circle, lines))


def _tangent_rails(circle, lines):
    if len(circle.ends) != 2 or not all(point and len(point) == 3 for point in circle.ends):
        return ()
    rails = []
    for point in circle.ends:
        radial = tuple(a - b for a, b in zip(point, circle.anchor.model))
        tangents = []
        for line in lines:
            a, b = (anchor.model for anchor in line.ends)
            if min(dist(point, a), dist(point, b)) > TOLERANCE or dist(a, b) < .002:
                continue
            cross = abs(sum((y - x) * r for x, y, r in zip(a, b, radial)))
            if cross <= .025 * dist(a, b) * circle.radius:
                tangents.append(line)
        if not tangents:
            return ()
        rails.append(max(tangents, key=lambda line: dist(*(a.model for a in line.ends))))
    return tuple(rails)


def _bend_profile_tasks(circles, lines, normal):
    groups = defaultdict(list)
    for circle in circles:
        if circle.axis == normal and .22 <= circle.fraction <= .28:
            groups[_point_key(circle.anchor.model)].append(circle)
    for group in groups.values():
        inner = min(group, key=lambda circle: circle.radius)
        if max(circle.radius for circle in group) - inner.radius <= TOLERANCE:
            continue
        rails = _tangent_rails(inner, lines)
        if len(rails) != 2:
            continue
        directions = [tuple(b-a for a, b in zip(*(point.model for point in line.ends))) for line in rails]
        lengths = [dist((0., 0., 0.), delta) for delta in directions]
        cosine = abs(sum(a*b for a, b in zip(*directions))) / (lengths[0] * lengths[1])
        if cosine > .05 or min(lengths) < inner.radius:
            continue
        anchors = tuple(Anchor(line.entity, line.ends[0].model,
                               tuple((a+b)/2 for a, b in zip(*(point.sheet for point in line.ends))))
                        for line in rails)
        # Place the angle on the inside of the bend, between the two legs.
        tangent_points = [min(line.ends, key=lambda a: dist(a.model, point)).sheet
                          for line, point in zip(rails, inner.ends)]
        corner = tuple(a+b-c for a, b, c in zip(*tangent_points, inner.anchor.sheet))
        rays = [max(line.ends, key=lambda a: dist(a.sheet, corner)).sheet for line in rails]
        unit = [tuple((b-a)/dist(point, corner) for a, b in zip(corner, point)) for point in rays]
        reach = min(.026, .85*min(dist(point, corner) for point in rays)) / sqrt(2)
        position = tuple(corner[i]+reach*(unit[0][i]+unit[1][i]) for i in (0, 1))
        yield Task((inner.anchor,), RADIUS, inner.radius, 0)
        yield Task(anchors, ANGLE, acos(cosine), 0, position=position)


def _folded_flange_tasks(vertices, circles, lines, model, normal, matrix, contour, profile_axis):
    if contour is None:
        return ()
    extents = list(_blank_extent_tasks(vertices, matrix, contour, circles))
    if normal == profile_axis:
        bend = list(_bend_profile_tasks(circles, lines, normal))
        if not bend:
            return [*extents, *_profile_tasks(lines, matrix),
                    *_profile_radii(circles, lines, vertices, normal, matrix)]
        # A bend's concentric arcs give the stock thickness. Dimension it
        # once on a straight flange, using the actual boundary vertices.
        for radius in (task for task in bend if task.kind == RADIUS):
            center = radius.targets[0].model
            thickness = max(c.radius for c in circles if c.axis == normal and dist(c.anchor.model, center) <= TOLERANCE)
            thickness -= radius.nominal
            candidates = [task for task in _profile_tasks(lines, matrix)
                          if abs(task.nominal-thickness) <= TOLERANCE]
            if candidates:
                bend.append(max(candidates, key=lambda task: task.direction))
                break
        return [*extents, *bend]
    cap = _rounded_end(circles, normal, contour, matrix[12])
    holes = list(_hole_tasks(circles, lines, model, normal, matrix, cap=cap, vertices=vertices))
    if cap is None:
        return [*extents, *holes, *_profile_radii(circles, lines, vertices, normal, matrix)]
    circle, along, _ = cap
    # The bend profile owns the leg lengths. Keep the width on the face
    # that also owns the hole row when both projections show the same stock.
    transverse = axis_of(matrix[1-along:9:3])
    counts = Counter(bore.axis for bore in model.bores)
    preferred = max((i for i in range(3) if i != profile_axis), key=lambda i: (counts[i], -i))
    show_width = transverse != profile_axis or normal == preferred
    return [*(task for task in extents if show_width and task.direction != along),
            Task((circle.anchor,), RADIUS, circle.radius, 0), *holes]


def _corner_cut_tasks(lines, matrix):
    """Project the cutbacks of an oblique edge that joins orthogonal rails."""
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE, 1e-6)
    rails = defaultdict(set)
    oblique = []
    for line in lines:
        a, b = line.ends
        differences = [abs(a.sheet[i] - b.sheet[i]) for i in range(3)]
        if differences[2] > tolerance:
            continue
        first, second = (
            tuple(round(x / tolerance) for x in point.sheet) for point in line.ends
        )
        x, y = differences[:2]
        if min(x, y) <= tolerance and max(x, y) >= .004 * scale:
            axis = int(y > x)
            rails[first].add(axis)
            rails[second].add(axis)
        elif min(x, y) >= .003 * scale:
            oblique.append((line, first, second, (x, y)))
    for line, first, second, components in oblique:
        if not any(1 - axis in rails[second] for axis in rails[first]):
            continue
        for axis, size in enumerate(components):
            yield Task(line.ends, LINEAR, size / scale, axis, projected=True)


def _aligned_width_tasks(lines, matrix):
    """Dimension two parallel boundaries that describe the same feature span."""
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE * 2, 1e-6)
    groups = defaultdict(list)
    for line in lines:
        first, second = line.ends
        dx = abs(first.sheet[0] - second.sheet[0])
        dy = abs(first.sheet[1] - second.sheet[1])
        along = 0 if dx >= dy else 1
        if (min(dx, dy) > tolerance or max(dx, dy) < .005 * scale
                or abs(first.sheet[2] - second.sheet[2]) > tolerance):
            continue
        start, end = sorted((first.sheet[along], second.sheet[along]))
        key = (along, *(round(x / tolerance) for x in (start, end, first.sheet[2])))
        point = tuple((a+b) / 2 for a, b in zip(first.sheet, second.sheet))
        groups[key].append(Anchor(line.entity, first.model, point, line.silhouette))
    for key, anchors in groups.items():
        direction = 1 - key[0]
        if len(anchors) != 2:
            continue
        a, b = anchors
        width = abs(b.sheet[direction] - a.sheet[direction]) / scale
        if width >= .003:
            yield Task((a, b), LINEAR, width, direction, projected=True)


def _stock_span_tasks(lines, matrix):
    """Measure an internal boundary from a coplanar rail spanning the stock."""
    if not lines:
        return
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE * 2, 1e-6)
    limits = [(min(a.sheet[i] for line in lines for a in line.ends),
               max(a.sheet[i] for line in lines for a in line.ends)) for i in (0, 1)]
    groups = defaultdict(list)
    for line in lines:
        a, b = line.ends
        dx, dy = (abs(a.sheet[i] - b.sheet[i]) for i in (0, 1))
        along = int(dy > dx)
        measure = 1 - along
        level = a.sheet[measure]
        lo, hi = limits[measure]
        if (min(dx, dy) > tolerance or max(dx, dy) < .008 * scale
                or abs(a.sheet[2] - b.sheet[2]) > tolerance
                or not lo + tolerance < level < hi - tolerance):
            continue
        start, stop = sorted((a.sheet[along], b.sheet[along]))
        midpoint = tuple((x + y) / 2 for x, y in zip(a.sheet, b.sheet))
        anchor = Anchor(line.entity, a.model, midpoint, line.silhouette)
        groups[along, round(a.sheet[2] / tolerance)].append((start, stop, level, anchor))

    for (along, _), rails in groups.items():
        lo, hi = limits[along]
        full = [rail for rail in rails if abs(rail[0] - lo) <= tolerance and abs(rail[1] - hi) <= tolerance]
        for start, stop, level, anchor in rails:
            at_low = abs(start - lo) <= tolerance
            at_high = abs(stop - hi) <= tolerance
            if at_low == at_high:
                continue
            for _, _, base_level, datum in full:
                width = abs(base_level - level) / scale
                if width >= .004 and (stop-start) / scale >= .15 * width:
                    yield Task((datum, anchor), LINEAR, width, 1 - along, projected=True)


def _recess_tasks(lines, matrix):
    """Dimension a symmetric recess bounded by a floor, two walls and rims."""
    scale = matrix[12]
    tolerance = max(scale * TOLERANCE, 1e-6)
    joints = defaultdict(list)
    planar = {}
    for line in lines:
        a, b = line.ends
        if abs(a.sheet[2] - b.sheet[2]) > tolerance:
            continue
        ends = tuple(tuple(round(x / tolerance) for x in point.sheet) for point in (a, b))
        key = tuple(sorted(ends))
        if ends[0] == ends[1] or key in planar:
            continue
        planar[key] = line
        for index, endpoint in enumerate(ends):
            joints[endpoint].append((key, line.ends[1 - index]))

    for key, floor in planar.items():
        a, b = floor.ends
        delta = [abs(a.sheet[i] - b.sheet[i]) for i in (0, 1)]
        along = int(delta[1] > delta[0])
        depth_axis = 1 - along
        if delta[along] < .008 * scale or delta[depth_axis] > tolerance:
            continue
        sides = []
        for sign, bottom in zip((-1, 1), sorted(floor.ends, key=lambda p: p.sheet[along])):
            node = tuple(round(x / tolerance) for x in bottom.sheet)
            options = []
            for wall_key, rim in joints[node]:
                if wall_key == key:
                    continue
                depth = rim.sheet[depth_axis] - bottom.sheet[depth_axis]
                run = sign * (rim.sheet[along] - bottom.sheet[along])
                if abs(depth) < .003 * scale or run < -tolerance:
                    continue
                rim_node = tuple(round(x / tolerance) for x in rim.sheet)
                if any(other != wall_key
                       and abs(end.sheet[depth_axis] - rim.sheet[depth_axis]) <= tolerance
                       and sign * (end.sheet[along] - rim.sheet[along]) >= .004 * scale
                       for other, end in joints[rim_node]):
                    options.append((depth, run, bottom, rim, planar[wall_key]))
            if len(options) != 1:
                break
            sides.extend(options)
        if len(sides) != 2 or any(abs(a - b) > tolerance for a, b in zip(sides[0][:2], sides[1][:2])):
            continue
        depth, run, bottom, rim, first_wall = sides[0]
        yield Task((bottom, rim), LINEAR, abs(depth) / scale, depth_axis, projected=True)
        if run <= tolerance:
            continue
        yield Task(floor.ends, LINEAR, delta[along] / scale, along)
        angle = acos((depth*depth - run*run) / (depth*depth + run*run))
        if .20 < angle < 2.95:
            walls = (first_wall, sides[1][4])
            anchors = tuple(Anchor(line.entity, line.ends[0].model,
                                   tuple((a + b) / 2 for a, b in zip(*(p.sheet for p in line.ends))))
                            for line in walls)
            yield Task(anchors, ANGLE, angle, along)


def _section_frame(vertices, model, normal, matrix, box):
    """Register the temporary section body, which COM can return translated."""
    axis = primary_axis(model)
    plan = choose_section(model, axis)
    identity = (0., 0., 0.)
    if not vertices or plan is None or normal in (axis, plan.along):
        return matrix, identity
    levels = Counter(round(v.model[normal] / TOLERANCE) for v in vertices)
    level, count = levels.most_common(1)[0]
    if count < max(4, len(vertices) / 2):
        return matrix, identity
    plane = [v for v in vertices if round(v.model[normal] / TOLERANCE) == level]
    shift = [0., 0., 0.]
    shift[normal] = sum(v.model[normal] for v in plane) / len(plane) - plan.coordinate
    for i in (axis, plan.along):
        low = min(v.model[i] for v in plane)
        high = max(v.model[i] for v in plane)
        if abs(high - low - model.spans[i]) > max(TOLERANCE * 2, model.spans[i] * .001):
            # A section that does not span the body's bounds has no reliable
            # landmark for recovering an in-plane translation here.
            return matrix, identity
        shift[i] = low - model.bounds[i]

    def midpoint(operation):
        projected = [operation(v.model) for v in plane]
        return tuple((min(p[i] for p in projected) + max(p[i] for p in projected)) / 2
                     for i in (0, 1))

    center = ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2)
    raw_center = midpoint(lambda point: project(point, matrix))
    original_center = midpoint(lambda point: project(tuple(a-b for a, b in zip(point, shift)), matrix))
    if dist(original_center, center) + .001 < dist(raw_center, center):
        # The IView transform maps original coordinates, while the returned
        # entities use the translated section body. Correct the translation.
        adjusted = list(matrix)
        for i in range(3):
            adjusted[9+i] -= matrix[12] * sum(shift[j] * matrix[3*j+i] for j in range(3))
        matrix = tuple(adjusted)
    return matrix, tuple(shift)


def _section_datums(lines, depth, scale):
    """Find transverse faces and vertices whose axial walls stop at an edge break."""
    tolerance = max(TOLERANCE*scale, 1e-6)
    planes = defaultdict(list)
    diagonals = []
    for first, second in (line.ends for line in lines):
        dx, dy, dz = (abs(a-b) for a, b in zip(first.sheet, second.sheet))
        if dz > tolerance or (dx, dy)[1-depth] <= tolerance:
            continue
        if (dx, dy)[depth] <= tolerance:
            planes[_point_key(first.model)].append((first, second))
            planes[_point_key(second.model)].append((second, first))
        else:
            diagonals.append((first, second))
    trimmed = defaultdict(list)
    for first, second in diagonals:
        for wall, face in ((first, second), (second, first)):
            trimmed[_point_key(wall.model)].extend(planes.get(_point_key(face.model), ()))
    return planes, trimmed


def _section_depth_task(pair, center, direction, scale, datums):
    depth = 1-direction
    planes, trimmed = datums
    tolerance = max(TOLERANCE*scale, 1e-6)
    keys = [_point_key(anchor.model) for anchor in pair]
    raw = abs(pair[1].sheet[depth]-pair[0].sheet[depth])
    if not any(key in trimmed and key not in planes for key in keys):
        return Task(pair, LINEAR, raw/scale, depth)

    # The end at a face opening continues away from the bore centre. The
    # shoulder at the bottom of the larger bore continues towards it.
    radial_sign = 1 if pair[0].sheet[direction] > center else -1
    faces = []
    for anchor, other, key in zip(pair, reversed(pair), keys):
        outward = 1 if anchor.sheet[depth] > other.sheet[depth] else -1
        candidates = [(near, radial_sign*(far.sheet[direction]-near.sheet[direction]) > 0.)
                      for near, far in planes.get(key) or trimmed.get(key, ())
                      if outward*(near.sheet[depth]-anchor.sheet[depth]) >= -tolerance]
        if not candidates:
            return None
        datum, entry = candidates[0]
        if any(role != entry or abs(point.sheet[depth]-datum.sheet[depth]) > tolerance
               for point, role in candidates[1:]):
            return None
        faces.append((datum, entry))
    (first, first_entry), (second, second_entry) = faces
    if first_entry == second_entry:
        # Two openings are a through bore, or the remaining wall under a
        # counterbore. Leave it to the stock and counterbore dimensions.
        return None
    return Task((first, second), LINEAR, abs(second.sheet[depth]-first.sheet[depth])/scale,
                depth, projected=True)


def _section_tasks(vertices, lines, model, normal, matrix, shift=(0., 0., 0.), *, datums):
    axis = primary_axis(model)
    section = choose_section(model, axis)
    if section is None or normal in (axis, section.along):
        return
    row = [abs(matrix[3 * section.along + i]) for i in (0, 1)]
    direction = row.index(max(row))
    if row[direction] < .9:
        return
    tolerance = max(matrix[12] * TOLERANCE, 1e-6)
    for original in section.bores:
        transverse = [i for i in range(3) if i != axis]
        bore = replace(original, center=tuple(c + shift[i] for c, i in zip(original.center, transverse)),
                       low=original.low + shift[axis], high=original.high + shift[axis])
        # Measure the span between actual edges of the cut at one axial
        # station. Restrictions keep independent or dashed geometry out.
        center = project(bore.point((bore.low + bore.high) / 2), matrix)[direction]
        half_width = bore.radius * matrix[12]
        ends = [v for v in vertices if bore.low - TOLERANCE <= v.model[axis] <= bore.high + TOLERANCE]
        pair = _pair_on_axis(ends, center - half_width, center + half_width, direction, tolerance)
        wall_pair = _section_walls(lines, bore, center, half_width, direction, tolerance)
        if wall_pair:
            yield Task(wall_pair, LINEAR, 2 * bore.radius, direction, projected=True, bore=original)
        if pair:
            yield Task(pair, LINEAR, 2 * bore.radius, direction, projected=True, bore=original)
        if not wall_pair and not pair:
            # Some builds omit cut entities. Predict wall hit points, then
            # validate the selected edge at each point before dimensioning.
            point = bore.point((bore.low + bore.high) / 2)
            middle = project(point, matrix)
            targets = []
            for sign in (-1, 1):
                sheet_point = list(middle)
                sheet_point[direction] += sign * half_width
                targets.append(Anchor(None, point, tuple(sheet_point)))
            yield Task(tuple(targets), LINEAR, 2 * bore.radius, direction, projected=True, bore=original)
        depths = sorted(
            project(bore.point(z), matrix)[1 - direction] for z in (bore.low, bore.high)
        )
        wall = [v for v in vertices if abs(abs(v.sheet[direction] - center) - half_width) <= tolerance]
        pair = _pair_on_axis(wall, *depths, 1 - direction, tolerance)
        if pair and (depth := _section_depth_task(pair, center, direction, matrix[12], datums)):
            yield depth


def _section_walls(lines, bore, center, half_width, direction, tolerance):
    walls = [[], []]
    for line in lines:
        a, b = line.ends
        mid = tuple((x + y) / 2 for x, y in zip(a.sheet, b.sheet))
        low, high = sorted((a.model[bore.axis], b.model[bore.axis]))
        overlap = min(high, bore.high) - max(low, bore.low)
        if (overlap < (bore.high - bore.low) * .45
                or abs(a.sheet[direction] - b.sheet[direction]) > tolerance):
            continue
        for index, sign in enumerate((-1, 1)):
            if abs(mid[direction] - center - sign * half_width) <= tolerance:
                walls[index].append((overlap, Anchor(line.entity, a.model, mid, line.silhouette)))
    if all(walls):
        return tuple(max(side, key=lambda item: item[0])[1] for side in walls)
    return None


def _through_section_openings(boxes, axis, along):
    ordered = sorted(boxes, key=lambda box: box[along])
    for left, right in zip(ordered, ordered[1:]):
        start, end = left[along+3], right[along]
        if end - start <= 2*TOLERANCE:
            continue
        if any(box[along] < end-TOLERANCE and box[along+3] > start+TOLERANCE
               for box in ordered):
            continue
        low, high = left[axis], left[axis+3]
        if max(abs(low-right[axis]), abs(high-right[axis+3])) <= 2*TOLERANCE:
            yield (start+end)/2, low, high


def _closed_inner_contours(strokes, bounds):
    resolution = .0001  # sheet coordinates, including tessellated arcs
    neighbors = defaultdict(set)
    for line in strokes:
        a, b = (tuple(round(value / resolution) for value in p) for p in line)
        if a != b:
            neighbors[a].add(b)
            neighbors[b].add(a)
    remaining = set(neighbors)
    while remaining:
        stack = [remaining.pop()]
        component = set(stack)
        while stack:
            for point in neighbors[stack.pop()] - component:
                component.add(point)
                remaining.remove(point)
                stack.append(point)
        if not all(len(neighbors[p]) == 2 for p in component):
            continue
        if not all(bounds[i] + resolution < point[i]*resolution < bounds[i+2] - resolution
                   for point in component for i in (0, 1)):
            continue
        yield tuple((tuple(v*resolution for v in a), tuple(v*resolution for v in b))
                    for a in component for b in neighbors[a] if a < b)


def _inside_contour(point, contour):
    x, y = point[:2]
    crossings = sum((ay > y) != (by > y) and ax+(bx-ax)*(y-ay)/(by-ay) > x
                    for (ax, ay), (bx, by) in contour)
    return crossings % 2 == 1


def _section_axis_lines(section, matrix, shift, hatch_boxes, base_matrix=None, contours=()):
    axis = section.bores[0].axis
    along = section.along
    normal = next(i for i in range(3) if i not in (axis, along))
    depth = max((0, 1), key=lambda i: abs(matrix[3*axis+i]))
    if abs(matrix[3*axis+depth]) < .999:
        return ()

    def on_model(station, height):
        point = [0., 0., 0.]
        point[along], point[axis], point[normal] = station, height, section.coordinate
        return point

    def on_sheet(station, height):
        return project(tuple(a+b for a, b in zip(on_model(station, height), shift)), matrix)

    transverse = [i for i in range(3) if i != axis]
    known = defaultdict(list)
    for bore in section.bores:
        known[_point_key(bore.point(0.))].append(bore)
    extents = {}
    for group in known.values():
        station = group[0].center[transverse.index(along)]
        extents[station] = (min(b.low for b in group), max(b.high for b in group))
    for station, low, high in _through_section_openings(hatch_boxes, axis, along):
        matches = [known_station for known_station in extents
                   if abs(known_station-station) <= 2*TOLERANCE]
        if matches:
            extents[matches[0]] = (low, high)
            continue
        # A gap can also be an exterior notch. Only treat it as a slot
        # when its centre is inside a closed internal contour of the parent.
        if contours:
            point = project(on_model(station, (low+high)/2), base_matrix)
            if any(_inside_contour(point, contour) for contour in contours):
                extents[station] = (low, high)

    lines = []
    extension = .002
    for station, (low, high) in sorted(extents.items()):
        points = sorted((on_sheet(station, low), on_sheet(station, high)), key=lambda p: p[depth])
        a, b = [list(point[:2]) + [0.] for point in points]
        a[depth] -= extension
        b[depth] += extension
        lines.append((tuple(a), tuple(b)))
    return tuple(lines)


def _cut_face_boxes(view, normal, coordinate):
    for raw in _optional(view, "GetFaceHatches") or ():
        face = _optional(as_dispatch(raw), "Face")
        box = _optional(face, "GetBox") if face is not None else None
        if box and len(box) == 6 and max(abs(box[i]-coordinate)
                                         for i in (normal, normal+3)) <= 2*TOLERANCE:
            yield box


def _section_parent_contours(view):
    base = _optional(view, "GetBaseView")
    if base is None:
        return None, ()
    strokes = view_strokes(base, get_value(base, "GetOutline")[:4])
    if not strokes:
        return None, ()
    return transform_of(base), tuple(_closed_inner_contours(strokes, geometry_outline(strokes, None)))


def _label(task, own, boxes, text, frame, *, at=None):
    x0, y0, x1, y1 = own
    x, y = at if at is not None else (
        sum(a.sheet[i] for a in task.targets) / len(task.targets) for i in (0, 1))
    low_x, low_y, high_x, high_y = frame.drawable
    if task.position is not None and low_x < task.position[0] < high_x and low_y < task.position[1] < high_y:
        return task.position
    for offset in (.008, .014, .020, .026):
        if task.kind in (DIAMETER, RADIUS):
            positions = ((x1 + offset, y), (x0 - offset, y), (x, y1 + offset), (x, y0 - offset))
        elif task.direction == 0:
            positions = ((x, y0 - offset), (x, y1 + offset))
        else:
            positions = ((x1 + offset, y), (x0 - offset, y))
        for xx, yy in positions:
            if not (low_x < xx < high_x and low_y < yy < high_y):
                continue
            if any(a - .003 < xx < c + .003 and b - .003 < yy < d + .003 for a, b, c, d in boxes):
                continue
            if all(abs(xx - old_x) >= .012 or abs(yy - old_y) >= .005 for old_x, old_y in text):
                return xx, yy
    return None


def _create(doc, view, task, position, matrix):
    get_value(doc, "ClearSelection2", True)
    manager = get_value(doc, "SelectionManager")
    if not get_value(doc, "ActivateView", get_value(view, "GetName2")):
        return None
    section = get_value(view, "Type") == 2
    for index, target in enumerate(task.targets):
        if target.entity is None:
            selected = _select_section_edge(doc, view, manager, target, index, matrix)
        elif target.silhouette:
            selection = get_value(manager, "CreateSelectData")
            selection.View = view
            selected = _optional(target.entity, "Select2", index > 0, selection)
        else:
            selected = _optional(view, "SelectEntity", target.entity, index > 0)
            if not selected and section:
                selection = get_value(manager, "CreateSelectData")
                selection.View = view
                selected = _optional(target.entity, "Select4", index > 0, selection)
                if not selected:
                    vertex = _optional(target.entity, "GetPoint")
                    if vertex:
                        selected = _select_section_vertex(doc, view, manager, target, index)
                    else:
                        selected = _select_section_edge(doc, view, manager, target, index, matrix)
        if not selected:
            return None
    if get_value(manager, "GetSelectedObjectCount2", -1) != len(task.targets):
        return None
    if task.projected:
        method = "AddHorizontalDimension2" if task.direction == 0 else "AddVerticalDimension2"
    else:
        method = {DIAMETER: "AddDiameterDimension2", RADIUS: "AddRadialDimension2"}.get(
            task.kind, "AddDimension2")
    display = get_value(doc, method, *position, 0.)
    if display is None:
        return None
    dimension = _optional(display, "GetDimension2", 0)
    annotation = _optional(display, "GetAnnotation")
    if dimension is not None and task.arc_ends:
        ends = [i for i in (1, 2) if _optional(dimension, "GetArcEndCondition", i) in (1, 2, 3)]
        if len(ends) != task.arc_ends or any(_optional(dimension, "SetArcEndCondition", i, SW_ARC_END_FARTHEST) != 0
                                           for i in ends):
            _discard(doc, view, display, annotation)
            return None
        get_value(doc, "GraphicsRedraw2")
    value = _optional(dimension, "GetSystemValue2", "") if dimension is not None else None
    kind = _optional(display, "Type2")
    linked = _optional(annotation, "GetAttachedEntityCount3") if annotation is not None else None
    if (linked is not None and linked >= len(task.targets) and kind is not None
            and value is not None and _key(kind, value)[0] == task.kind
            and isclose(value, task.nominal, abs_tol=TOLERANCE, rel_tol=1e-5)):
        if task.overall and _optional(view, "IsFlatPatternView"):
            if _optional(display, "Diametric"):
                display.Diametric = False
            if _optional(display, "GetText", SW_DIMENSION_TEXT_PREFIX) in ("<MOD-DIAM>", "Ø", "⌀"):
                get_value(display, "SetText", SW_DIMENSION_TEXT_PREFIX, "")
        return annotation
    _discard(doc, view, display, annotation)
    return None


def _select_section_at(doc, view, manager, target, index, kind):
    x, y = target.sheet[:2]
    low_x, low_y, high_x, high_y = get_value(view, "GetOutline")[:4]
    if not (low_x < x < high_x and low_y < y < high_y):
        return None
    ext = get_value(doc, "Extension")
    if not _optional(ext, "SelectByID2", "", kind, x, y, 0., index > 0, 0,
                     VARIANT(VT_DISPATCH, None), 0):
        return None
    view_name = _optional(view, "GetName2")
    selected_view = _optional(manager, "GetSelectedObjectsDrawingView2", index + 1, -1)
    if selected_view is not None and _optional(selected_view, "GetName2") != view_name:
        return None
    return _optional(manager, "GetSelectedObject6", index + 1, -1)


def _select_section_vertex(doc, view, manager, target, index):
    vertex = _select_section_at(doc, view, manager, target, index, "VERTEX")
    point = _optional(vertex, "GetPoint") if vertex is not None else None
    return bool(point and len(point) == 3 and dist(point, target.model) <= TOLERANCE)


def _select_section_edge(doc, view, manager, target, index, matrix=None):
    x, y = target.sheet[:2]
    edge = _select_section_at(doc, view, manager, target, index, "EDGE")
    curve = _optional(edge, "GetCurve") if edge is not None else None
    if curve is None or not _optional(curve, "IsLine"):
        return False
    points = []
    for getter, fallback in (("GetStartVertex", "GetStartPoint"), ("GetEndVertex", "GetEndPoint")):
        vertex = _optional(edge, getter)
        if vertex is not None:
            points.append(_optional(vertex, "GetPoint"))
            continue
        point = _optional(edge, fallback)
        points.append(_optional(point, "ArrayData") if point is not None else None)
    if not all(p and len(p) >= 3 for p in points):
        return False
    if matrix is None:
        matrix = transform_of(view)
    a, b = (project(p, matrix)[:2] for p in points)
    delta = (b[0] - a[0], b[1] - a[1])
    length2 = sum(v * v for v in delta)
    if length2 < 1e-12:
        return False
    t = ((x - a[0]) * delta[0] + (y - a[1]) * delta[1]) / length2
    hit = (a[0] + t * delta[0], a[1] + t * delta[1])
    return 0 <= t <= 1 and dist(hit, (x, y)) <= max(matrix[12] * TOLERANCE * 2, 1e-6)


def _discard(doc, view, display, annotation):
    get_value(doc, "ClearSelection2", True)
    extension = get_value(doc, "Extension")
    null = VARIANT(VT_DISPATCH, None)
    name = _optional(display, "GetNameForSelection")
    selected = name and _optional(extension, "SelectByID2", name, "DIMENSION", 0., 0., 0.,
                                  False, 0, null, 0)
    if not selected and annotation is not None:
        selected = _optional(annotation, "Select3", False, null)
    if selected:
        _optional(extension, "DeleteSelection2", 0)
        if name and not any(_optional(item, "GetNameForSelection") == name
                            for item in (_optional(view, "GetDisplayDimensions") or ())):
            return
    if annotation is not None and _hide(annotation):
        return
    raise RuntimeError("Не удалось убрать размер без подтверждённой привязки")


def add_dimensions(doc, model, frame, *, flat_model=None):
    from .hole_annotations import _thread_centers, projected_holes

    if model is None:
        return 0, [], [], []
    views = [(view, bool(_optional(view, "IsFlatPatternView"))) for view in drawing_views(doc)]
    views.sort(key=lambda item: item[1])
    if any(flat for _, flat in views):
        get_value(doc, "ViewZoomtofit2")
    created_geometry = set()
    created_profiles = set()
    # These generated labels get their lanes first. Imported labels will be
    # moved around them by the native text layout after dimensioning.
    text = []
    boxes = [tuple(get_value(view, "GetOutline")[:4]) for view, _ in views]
    added = 0
    cut_diameters = []
    section_axes = []
    hole_views = []
    flange_profile = perpendicular_flange_profile(model)
    for (view, flat), box in zip(views, boxes):
        matrix = transform_of(view)
        normal = axis_of(matrix[2:9:3])
        if normal is None and not flat:
            continue
        vertices, circles, lines = _anchors(view, matrix)
        section = get_value(view, "Type") == 2
        if not section:
            threads = () if flat else _thread_centers(view)
            holes = projected_holes(circles, flat_model if flat else model, matrix, threads)
            if holes:
                hole_views.append((view, holes, matrix, flat))
        view_scope = get_value(view, "GetName2") if section or flat else None
        imported = _existing(view)
        created_in_view = set()
        profile_shift = (_folded_profile_shift(vertices, model, normal, _optional(view, "Position"))
                         if model.sheet_metal and view_scope is None else None)
        if flange_profile is not None or flat:
            shift = _arc_vertex_shift(vertices, circles, _optional(view, "Position"))
            if shift is None and not flat:
                shift = profile_shift
            if shift is not None:
                vertices, lines = _registered_vertices(vertices, lines, matrix, shift)
                profile_shift = None if flat else (0., 0., 0.)
        if section:
            matrix, shift = _section_frame(vertices, model, normal, matrix, box)
            def reproject(anchor):
                return replace(anchor, sheet=project(anchor.model, matrix))
            vertices = [reproject(vertex) for vertex in vertices]
            lines = [replace(line, ends=tuple(map(reproject, line.ends))) for line in lines]
        contour = geometry_outline(view_strokes(view, box), None) if flat or flange_profile is not None else None
        if flat:
            tasks = _blank_extent_tasks(vertices, matrix, contour, circles,
                                        lines if flange_profile is not None else ())
        elif flange_profile is not None:
            tasks = _folded_flange_tasks(vertices, circles, lines, model, normal, matrix, contour,
                                         flange_profile)
        else:
            tasks = list(_overall_tasks(vertices, matrix))
            if section:
                model_axis = primary_axis(model)
                depth_direction = max((0, 1), key=lambda i: abs(matrix[3*model_axis+i]))
                datums = _section_datums(lines, depth_direction, matrix[12])
                cut_lines = lines + list(_section_silhouettes(view, matrix))
                tasks.extend(_section_tasks(vertices, cut_lines, model, normal, matrix, shift, datums=datums))
                plan = choose_section(model, model_axis)
                if plan is not None and normal not in (model_axis, plan.along):
                    hatches = tuple(_cut_face_boxes(view, normal, plan.coordinate))
                    base_matrix, contours = _section_parent_contours(view) if hatches else (None, ())
                    axes = _section_axis_lines(plan, matrix, shift, hatches, base_matrix, contours)
                    section_axes.append((view, axes))
            else:
                tasks.extend(_recess_tasks(lines, matrix))
                tasks.extend(_stock_span_tasks(lines, matrix))
                tasks.extend(_hole_tasks(circles, lines, model, normal, matrix))
                tasks.extend(_profile_radii(circles, lines, vertices, normal, matrix))
            tasks.extend(_corner_cut_tasks(lines, matrix))
            for task in _profile_tasks(lines, matrix):
                if section and task.direction == depth_direction and any(
                        (point := _point_key(anchor.model)) in datums[1] and point not in datums[0]
                        for anchor in task.targets):
                    continue
                tasks.append(task)
            tasks.extend(_aligned_width_tasks(lines, matrix))
        profiles = {id(task): _physical_profile_key(task, matrix, profile_shift)
                    for task in tasks if task.profile} if profile_shift is not None else {}
        # If an edge is already dimensioned, omit parallel candidates at
        # the same stations in this view as well.
        covered_profile_spans = {_profile_interval(profile) for profile in profiles.values()
                                 if profile is not None and profile in created_profiles}
        for task in tasks:
            key = _key(task.kind, task.nominal)
            local = (key, task.direction if task.kind == LINEAR else None)
            geometry = _geometry_key(task, matrix, normal, view_scope)
            if key in imported or local in created_in_view or geometry in created_geometry:
                continue
            profile = profiles.get(id(task))
            if profile is not None and _profile_interval(profile) in covered_profile_spans:
                continue
            flat_center = ((box[0]+box[2])/2, (box[1]+box[3])/2) if flat and flange_profile is not None else None
            position = _label(task, box, boxes, text, frame, at=flat_center)
            if position is None:
                continue
            annotation = _create(doc, view, task, position, matrix)
            if annotation is not None:
                added += 1
                if task.bore is not None:
                    cut_diameters.append((task.bore, annotation))
                created_in_view.add(local)
                created_geometry.add(geometry)
                if profile is not None:
                    created_profiles.add(profile)
                text.append(position)
        get_value(doc, "ClearSelection2", True)
    return added, cut_diameters, section_axes, hole_views
