"""Arrange dimension labels against native text and view geometry."""
from collections import defaultdict, namedtuple
from dataclasses import dataclass, replace
from itertools import chain, islice
from math import cos, dist, hypot, isclose, isfinite, sin, tau
from pathlib import Path
from re import sub

from pythoncom import VT_BYREF, VT_VARIANT, com_error
from win32com.client import VARIANT

from .drawing_views import drawing_views
from .solidworks import as_dispatch, get_value


VisibleDimension = namedtuple("VisibleDimension", "view display annotation dimension kind")
CHAIN_TOLERANCE = .0001  # on the drawing sheet, in metres
NESTED_CLEARANCE = .006
INLINE_CLEARANCE = .0015
CALLOUT_RADIUS = .014
FLAT_PATTERN_CAPTION = "Развертка"


@dataclass(frozen=True)
class LinearSpan:
    axis: int
    start: float
    end: float
    baseline: float


def optional(obj, method, *args):
    try:
        return get_value(obj, method, *args)
    except (AttributeError, com_error):
        return None


def visible_dimensions(doc):
    for view in drawing_views(doc):
        for display in optional(view, "GetDisplayDimensions") or ():
            annotation = optional(display, "GetAnnotation")
            if annotation is not None and optional(annotation, "Visible") == 1:
                dimension = optional(display, "GetDimension2", 0)
                yield VisibleDimension(view, display, annotation, dimension, optional(display, "Type2"))


def set_black_dimension_color(doc):
    dimensions = list(visible_dimensions(doc))
    for item in dimensions:
        # Set an explicit drawing color so driven and imported sizes agree.
        item.annotation.Color = 0  # black COLORREF
    if dimensions:
        get_value(doc, "GraphicsRedraw2")


def text_extent(parts):
    corners = []
    for pos, width, height, angle in parts:
        if not pos or len(pos) < 2 or not all(isinstance(v, (int, float)) and isfinite(v)
                                               for v in (*pos[:2], width, height, angle)):
            continue
        if not (.00005 <= width <= .25 and .0002 <= height <= .10):
            continue
        c, s = cos(angle), sin(angle)
        # DisplayData positions are measured on the sheet; local X follows
        # the text baseline and local Y is perpendicular to it.
        for u in (0., width):
            for v in (-.0005, height):
                corners.append((pos[0] + u*c - v*s, pos[1] + u*s + v*c))
    if not corners:
        return None
    return tuple(f(p[i] for p in corners) for f in (min, max) for i in (0, 1))


def text_geometry(annotation, *, display_data=None):
    data = display_data if display_data is not None else optional(annotation, "GetDisplayData")
    if data is None:
        return None, None, ()
    parts = []
    text = []
    orientation = None
    for index in range(optional(data, "GetTextCount") or 0):
        value = optional(data, "GetTextAtIndex", index)
        width = optional(data, "GetTextInBoxWidthAtIndex", index)
        if not value or width is None:
            continue
        angle = optional(data, "GetTextAngleAtIndex", index)
        if orientation is None:
            orientation = angle
        text.append(str(value).strip())
        parts.append((optional(data, "GetTextPositionAtIndex", index), width,
                      optional(data, "GetTextHeightAtIndex", index), angle))
    return text_extent(parts), orientation, tuple(text)


def dimension_lines(annotation):
    data = optional(annotation, "GetDisplayData")
    if data is None:
        return ()
    strokes = []
    for i in range(optional(data, "GetLineCount") or 0):
        line = optional(data, "GetLineAtIndex", i)
        if line and len(line) == 7 and all(isfinite(v) for v in line):
            strokes.append((tuple(line[1:3]), tuple(line[4:6])))
    return tuple(strokes)


def _linear_span(item, *, display_data=None):
    if item.kind not in (2, 11, 12):
        return None
    data = display_data if display_data is not None else optional(item.annotation, "GetDisplayData")
    if data is None or optional(data, "GetArrowHeadCount") != 2:
        return None
    arrows = [optional(data, "GetArrowHeadAtIndex2", i) for i in range(2)]
    if not all(arrow and len(arrow) >= 6 and all(map(isfinite, arrow[:6])) for arrow in arrows):
        return None
    first, second = arrows
    dx, dy = abs(first[0]-second[0]), abs(first[1]-second[1])
    axis = int(dy > dx)
    across = 1-axis
    if (item.kind in (11, 12) and item.kind != 11+axis
            or abs(first[across]-second[across]) > CHAIN_TOLERANCE
            or any(abs(arrow[axis+3]) < .999 for arrow in arrows)):
        return None
    start, end = sorted((first[axis], second[axis]))
    if end-start <= CHAIN_TOLERANCE:
        return None
    return LinearSpan(axis, start, end, (first[across]+second[across])/2)


def _touching_chains(spans):
    chains = []
    for index, span in sorted(spans, key=lambda entry: (entry[1].start, entry[1].end)):
        matches = [chain for chain in chains if abs(chain[-1][1].end-span.start) <= CHAIN_TOLERANCE]
        if matches:
            min(matches, key=lambda chain: abs(chain[-1][1].baseline-span.baseline)).append((index, span))
        else:
            chains.append([(index, span)])
    return (chain for chain in chains if len(chain) > 1)


def _polyline_strokes(values, transform):
    if not values or not transform or len(transform) != 3 or transform[2] <= 0:
        return ()
    tx, ty, scale = transform
    strokes = []
    offset = 0
    while offset + 8 < len(values):
        # Type, geometry payload length, payload, six display attributes,
        # point count, then XYZ points in unscaled view coordinates.
        count_at = offset + 8 + int(values[offset + 1])
        if not offset + 8 <= count_at < len(values):
            return ()
        next_polyline = count_at + 1 + 3 * int(values[count_at])
        if not count_at < next_polyline <= len(values):
            return ()
        points = [(tx + scale*values[i], ty + scale*values[i+1])
                  for i in range(count_at + 1, next_polyline, 3)]
        strokes.extend(zip(points, points[1:]))
        offset = next_polyline
    return tuple(strokes) if offset == len(values) else ()


def view_strokes(view, outline):
    coordinates = VARIANT(VT_BYREF | VT_VARIANT, None)
    optional(view, "GetPolylines7", 0, coordinates)
    strokes = _polyline_strokes(coordinates.value, optional(view, "GetXform"))
    # Keep the external placement rule on views where the projected edges
    # are unavailable or use a coordinate system we cannot verify.
    if not all(outline[0] - .001 <= x <= outline[2] + .001
               and outline[1] - .001 <= y <= outline[3] + .001
               for line in strokes for x, y in line):
        return ()
    return strokes


def geometry_outline(strokes, fallback):
    if not strokes:
        return fallback
    points = [point for line in strokes for point in line]
    return tuple(f(point[i] for point in points) for f in (min, max) for i in (0, 1))


def intersects(a, b, clearance=0.):
    return (a[0] < b[2] + clearance and b[0] < a[2] + clearance
            and a[1] < b[3] + clearance and b[1] < a[3] + clearance)


def fits(box, allowed, views, used):
    return (box is not None and allowed[0] + .001 <= box[0]
            and allowed[1] + .001 <= box[1]
            and box[2] <= allowed[2] - .001 and box[3] <= allowed[3] - .001
            and not any(intersects(box, view, .0008) for view in views)
            and not any(intersects(box, previous, .0012) for previous in used))


def crosses(segment, box):
    start, end = segment
    first, last = 0., 1.
    for axis in (0, 1):
        low, high = box[axis] + .00015, box[axis+2] - .00015
        delta = end[axis] - start[axis]
        if abs(delta) < 1e-10:
            if not low < start[axis] < high:
                return False
            continue
        enters, leaves = sorted(((low-start[axis])/delta, (high-start[axis])/delta))
        first, last = max(first, enters), min(last, leaves)
        if first >= last:
            return False
    return True


def _leader_hits(segment, obstacle, center, radius):
    """Test crossings outside the circular edge to which the leader attaches."""
    a, b = segment
    c, d = obstacle
    if any(max(a[i], b[i]) < min(c[i], d[i]) or max(c[i], d[i]) < min(a[i], b[i]) for i in (0, 1)):
        return False
    x, y = b[0]-a[0], b[1]-a[1]
    u, v = d[0]-c[0], d[1]-c[1]
    px, py = c[0]-a[0], c[1]-a[1]
    det = x*v-y*u
    if abs(det) < 1e-12:
        if abs(px*y-py*x) > 1e-12:
            return False
        along = 0 if abs(x) >= abs(y) else 1
        span = b[along]-a[along]
        if abs(span) < 1e-6:
            return False
        low, high = sorted(( (c[along]-a[along])/span, (d[along]-a[along])/span ))
        left, right = max(0., low), min(1., high)
        return left <= right and any(
            dist((a[0]+t*x, a[1]+t*y), center) > radius for t in (left, right))
    t, s = (px*v-py*u)/det, (px*y-py*x)/det
    return 0. <= t <= 1. and 0. <= s <= 1. and dist((a[0]+t*x, a[1]+t*y), center) > radius


def _clear_leader(strokes, obstacles, center, radius, *, border=()):
    if not strokes:
        return False
    hits = ((line, edge) for line in strokes for edge in obstacles if _leader_hits(line, edge, center, radius))
    first = next(hits, None)
    if first is None:
        return True
    line, edge = first
    return edge in border and _line_crossings((line,), (edge,)) == 1 and next(hits, None) is None


def _leader_positions(anchors, width, height, allowed, own, foreign, used, obstacles, *, border=(), refine=False):
    proposals = []
    diagonal = .7071067812
    directions = ((1., 0.), (-1., 0.), (0., 1.), (0., -1.), (diagonal, diagonal),
                  (-diagonal, diagonal), (diagonal, -diagonal), (-diagonal, -diagonal))
    reaches = (.006, .010, .015, .021, .029, .040)
    if refine:
        directions = tuple((cos(step*tau/48), sin(step*tau/48)) for step in range(48))
        reaches = (.004, .006, .008, .010, .012, .015, .018, .021, .024, .029, .034, .040)
    for anchor, center, radius in anchors:
        for reach in reaches:
            for dx, dy in directions:
                tip = (center[0]+dx*radius, center[1]+dy*radius)
                elbow = (tip[0]+dx*reach, tip[1]+dy*reach)
                side = 1 if dx >= 0 else -1
                x = elbow[0] + side*width/2
                y = elbow[1]
                box = (x-width/2, y-height/2, x+width/2, y+height/2)
                if not fits(box, allowed, foreign, used):
                    continue
                if any(crosses(edge, box) for edge in obstacles):
                    continue
                outside = not intersects(box, own)
                if not _clear_leader(((tip, elbow),), obstacles, center, radius+.0004,
                                     border=border if outside else ()):
                    continue
                if any(crosses((tip, elbow), used_box) for used_box in used):
                    continue
                proposals.append((outside, reach, anchor, (x, y), center, radius))
    return sorted(proposals, key=lambda row: row[:2])


def _fits_detail(box, allowed, other_boxes, used, own, model_strokes):
    if not fits(box, allowed, other_boxes, used):
        return False
    if not intersects(box, own, .0008):
        return True
    margin = .001
    padded = (box[0]-margin, box[1]-margin, box[2]+margin, box[3]+margin)
    return (bool(model_strokes) and own[0] <= padded[0] and own[1] <= padded[1]
            and padded[2] <= own[2] and padded[3] <= own[3]
            and not any(crosses(line, padded) for line in model_strokes))


def _routes_fit(box, strokes, foreign_views, occupied, *, text_clearance=0.):
    if text_clearance:
        x0, y0, x1, y1 = box
        box = (x0-text_clearance, y0-text_clearance, x1+text_clearance, y1+text_clearance)
    for text, lines in occupied:
        if any(crosses(new, text) for new in strokes) or any(crosses(old, box) for old in lines):
            return False
    return not any(crosses(stroke, other) for other in foreign_views for stroke in strokes)


def format_zones(doc, frame):
    sheet = get_value(doc, "GetCurrentSheet")
    template = str(optional(sheet, "GetTemplateName") or "")
    if Path(template).stem.casefold().endswith(" - gost_sh1"):
        # Stock GOST formats put the 70 x 14 mm document designation in the
        # top left of the frame. Keep the cell clear even when it is empty.
        return ((.020, frame.height - .019, .090, frame.height - .005),)
    sheet_view = optional(doc, "GetFirstView")
    zones = []
    if sheet_view is not None:
        for note in optional(sheet_view, "GetNotes") or ():
            ann = optional(as_dispatch(note), "GetAnnotation")
            if ann is not None:
                box, _, _ = text_geometry(ann)
                if box is not None:
                    zones.append(box)
    return tuple(zones)


def _section_label_notes(view):
    if get_value(view, "Type") != 2:
        return
    for raw in optional(view, "GetNotes") or ():
        note = as_dispatch(raw)
        linked = str(optional(note, "PropertyLinkedText") or "")
        if "<VLLABEL>" in linked:
            yield note, linked


def section_captions(view):
    for note, _ in _section_label_notes(view):
        annotation = optional(note, "GetAnnotation")
        if annotation is not None and optional(annotation, "Visible") == 1:
            yield annotation


def omit_repeated_section_scales(doc):
    sheet = get_value(doc, "GetCurrentSheet")
    properties = get_value(sheet, "GetProperties2")
    sheet_scale = properties[2] / properties[3]
    for view in drawing_views(doc):
        if optional(view, "Type") != 2 or not isclose(get_value(view, "ScaleDecimal"), sheet_scale):
            continue
        for note, linked in _section_label_notes(view):
            label = sub(r"\s*\((?:\s*<VLSCALEV?>)+\s*\)|(?:\s*<VLSCALEV?>)+", "", linked)
            if label != linked:
                note.PropertyLinkedText = label


def section_line_zones(view):
    for raw in optional(view, "GetSectionLines") or ():
        section = as_dispatch(raw)
        label = str(optional(section, "GetLabel") or "")
        text = optional(section, "GetTextInfo") or ()
        arrows = optional(section, "GetArrowInfo") or ()
        fmt = optional(section, "GetTextFormat")
        height = optional(fmt, "CharHeight") if fmt is not None else None
        if not height or not .0002 <= height <= .05:
            continue
        # These coordinates and the character height are measured on the sheet.
        if label and len(text) == 6 and all(map(isfinite, text)):
            width = len(label)*height
            for x, y in (text[:2], text[3:5]):
                yield x-.001, y-height-.001, x+width+.001, y+.001
        if len(arrows) == 12 and all(map(isfinite, arrows)):
            pad = height/4
            for i in (0, 6):
                x0, y0, x1, y1 = arrows[i], arrows[i+1], arrows[i+3], arrows[i+4]
                yield min(x0,x1)-pad, min(y0,y1)-pad, max(x0,x1)+pad, max(y0,y1)+pad


def _place_caption(annotation, own, allowed, outlines, occupied):
    box, _, _ = text_geometry(annotation)
    if box is None or fits(box, allowed, outlines, occupied):
        return box
    position = optional(annotation, "GetPosition")
    if not position or len(position) < 3:
        return box
    width, height = box[2] - box[0], box[3] - box[1]
    x = (own[0] + own[2] - width) / 2
    for y in (own[3] + .002, own[1] - height - .002):
        candidate = (x, y, x + width, y + height)
        if not fits(candidate, allowed, outlines, occupied):
            continue
        optional(annotation, "SetPosition2", position[0] + x - box[0], position[1] + y - box[1], position[2])
        placed, _, _ = text_geometry(annotation)
        if fits(placed, allowed, outlines, occupied):
            return placed
    optional(annotation, "SetPosition2", *position[:3])
    return box


def _place_flat_caption(annotation, outline):
    position = optional(annotation, "GetPosition")
    if not position or len(position) < 3:
        return None
    x = (outline[0] + outline[2]) / 2
    y = outline[3] + .007
    if optional(annotation, "SetPosition2", x, y, position[2]) is False:
        return None
    box, _, _ = text_geometry(annotation)
    if box is None:
        return None
    dx = x - (box[0] + box[2]) / 2
    dy = outline[3] + .004 - box[1]
    optional(annotation, "SetPosition2", x+dx, y+dy, position[2])
    return text_geometry(annotation)[0]


def candidates(anchor, box, view, kind, angle):
    x, y, z = anchor
    x0, y0, x1, y1 = view
    label_x, label_y = (box[0] + box[2])/2, (box[1] + box[3])/2
    half_w, half_h = (box[2]-box[0])/2, (box[3]-box[1])/2
    horizontal = kind == 11 or (kind != 12 and (angle is None or abs(cos(angle)) >= .7))
    radial = kind in (3, 5, 6, 10)
    cx, cy = min(max(label_x, x0), x1), min(max(label_y, y0), y1)
    for delta in (.004, .010, .016, .022, .030):
        choices = []
        for jog in (0., -.007, .007):
            if horizontal or radial:
                choices.extend(((cx + jog, y0-half_h-delta), (cx + jog, y1+half_h+delta)))
            if not horizontal or radial:
                choices.extend(((x1+half_w+delta, cy + jog), (x0-half_w-delta, cy + jog)))
        for xx, yy in sorted(choices, key=lambda p: (p[0]-label_x)**2+(p[1]-label_y)**2):
            dx, dy = xx-label_x, yy-label_y
            yield (x+dx, y+dy, z), (box[0]+dx, box[1]+dy, box[2]+dx, box[3]+dy)


def _interior_candidates(anchor, box, view):
    center = tuple((box[i]+box[i+2])/2 for i in (0, 1))
    room = tuple((box[i+2]-box[i])/2 + .002 for i in (0, 1))
    if any(view[i] + 2*room[i] > view[i+2] for i in (0, 1)):
        return
    x, y = (min(max(center[i], view[i]+room[i]), view[i+2]-room[i]) for i in (0, 1))
    for dx, dy in ((0., 0.), (.005, 0.), (-.005, 0.), (0., .005), (0., -.005),
                   (.010, 0.), (-.010, 0.), (0., .010), (0., -.010)):
        sx, sy = x+dx-center[0], y+dy-center[1]
        if abs(sx)+abs(sy) <= .042:
            yield (anchor[0]+sx, anchor[1]+sy, anchor[2]), (
                box[0]+sx, box[1]+sy, box[2]+sx, box[3]+sy)


def _hide(ann):
    ann.Visible = 3
    if optional(ann, "Visible") != 3:
        raise RuntimeError("SolidWorks не скрыл размер без свободного места")


def _try_chain(members, lane, allowed, outline, foreign, occupied):
    original = [optional(item.annotation, "GetPosition") for item, _ in members]
    if not all(position and len(position) == 3 for position in original):
        return None
    moved = []
    placements = []
    for (item, span), position in zip(members, original):
        delta = lane-span.baseline
        if abs(delta) > CHAIN_TOLERANCE:
            target = list(position)
            target[1-span.axis] += delta
            if not optional(item.annotation, "SetPosition2", *target):
                break
            moved.append((item.annotation, position))
        updated = _linear_span(item)
        if (updated is None or updated.axis != span.axis
                or max(abs(updated.start-span.start), abs(updated.end-span.end),
                       abs(updated.baseline-lane)) > CHAIN_TOLERANCE):
            break
        box, _, _ = text_geometry(item.annotation)
        strokes = dimension_lines(item.annotation)
        used = [label for label, _ in occupied + placements]
        if (not fits(box, allowed, (outline, *foreign), used)
                or not _routes_fit(box, strokes, foreign, occupied + placements)):
            break
        placements.append((box, strokes))
    else:
        return placements, len(moved)
    for annotation, position in reversed(moved):
        if not optional(annotation, "SetPosition2", *position):
            raise RuntimeError("SolidWorks не восстановил положение размера")
    return None


def _encloses(outer, inner):
    return (outer.start <= inner.start + CHAIN_TOLERANCE
            and outer.end >= inner.end - CHAIN_TOLERANCE
            and outer.end-outer.start > inner.end-inner.start + CHAIN_TOLERANCE)


def _separate_nested(spans, side, placed, occupied, allowed, outline, foreign):
    lanes = defaultdict(list)
    for index, span in spans.items():
        lanes[round(span.baseline/CHAIN_TOLERANCE)].append((index, span))
    rows = [tuple(index for index, _ in chain) for lane in lanes.values()
            for chain in _touching_chains(lane)]
    connected = {index for row in rows for index in row}
    rows.extend((index,) for index in spans if index not in connected)
    rows.sort(key=lambda row: max(spans[index].end-spans[index].start for index in row))
    moved = 0
    for row in rows:
        members = [(placed[index][0], spans[index]) for index in row]
        children = [other for index, other in spans.items() if index not in row
                    and any(_encloses(span, other) for _, span in members)]
        if not children:
            continue
        minimum = max(side*child.baseline for child in children) + NESTED_CLEARANCE
        if min(side*span.baseline for _, span in members) >= minimum - CHAIN_TOLERANCE:
            continue
        others = [placement for index, placement in enumerate(occupied) if index not in row]
        for extra in (0., .003, .006):
            lane = side*(minimum + extra)
            if any(index not in row and abs(other.baseline-lane) <= CHAIN_TOLERANCE
                   and any(min(span.end, other.end)-max(span.start, other.start) > CHAIN_TOLERANCE
                           for _, span in members) for index, other in spans.items()):
                continue
            result = _try_chain(members, lane, allowed, outline, foreign, others)
            if result is not None:
                placements, count = result
                for index, placement in zip(row, placements):
                    occupied[index] = placement
                    spans[index] = replace(spans[index], baseline=lane)
                moved += count
                break
    return moved


def _align_chains(placed, occupied, outlines, reserved, allowed):
    groups = defaultdict(dict)
    for index, (_, name, span) in enumerate(placed):
        if name not in outlines or span is None:
            continue
        across = 1-span.axis
        outline = outlines[name]
        if span.baseline < outline[across] - .001:
            groups[name, span.axis, -1][index] = span
        elif span.baseline > outline[across+2] + .001:
            groups[name, span.axis, 1][index] = span
    moved = 0
    for (name, _, side), spans in groups.items():
        foreign = (*reserved, *(outline for owner, outline in outlines.items() if owner != name))
        moved += _separate_nested(spans, side, placed, occupied, allowed, outlines[name], foreign)
        for chain in _touching_chains(spans.items()):
            indices = {index for index, _ in chain}
            lanes = {span.baseline for _, span in chain}
            if max(lanes)-min(lanes) <= CHAIN_TOLERANCE:
                continue
            start, end = chain[0][1].start, chain[-1][1].end
            other_spans = [span for index, span in spans.items() if index not in indices
                           and min(end, span.end)-max(start, span.start) > CHAIN_TOLERANCE]
            others = [row for index, row in enumerate(occupied) if index not in indices]
            members = [(placed[index][0], span) for index, span in chain]
            for lane in sorted(lanes, key=lambda coordinate: side*coordinate):
                if any(abs(other.baseline-lane) <= CHAIN_TOLERANCE for other in other_spans):
                    continue
                if any((_encloses(other, span) and side*(other.baseline-lane) < NESTED_CLEARANCE-CHAIN_TOLERANCE)
                       or (_encloses(span, other) and side*(lane-other.baseline) < NESTED_CLEARANCE-CHAIN_TOLERANCE)
                       for _, span in chain for other in other_spans):
                    continue
                result = _try_chain(members, lane, allowed, outlines[name], foreign, others)
                if result is None:
                    continue
                placements, count = result
                for (index, span), placement in zip(chain, placements):
                    occupied[index] = placement
                    spans[index] = replace(span, baseline=lane)
                moved += count
                break
    return moved


def _translated_box(box, dx, dy):
    return box[0]+dx, box[1]+dy, box[2]+dx, box[3]+dy


def _span_overhang(span, strokes):
    axis, across = span.axis, 1-span.axis
    return max((max(span.start-min(a[axis], b[axis]), max(a[axis], b[axis])-span.end, 0.)
                for a, b in strokes if max(abs(a[across]-span.baseline), abs(b[across]-span.baseline))
                <= CHAIN_TOLERANCE), default=0.)


def _center_linear_labels(placed, occupied, outlines, edges, reserved, allowed):
    moved = 0
    for index, (item, name, span) in enumerate(placed):
        if span is None:
            continue
        box, strokes = occupied[index]
        axis = span.axis
        center = (box[axis]+box[axis+2])/2
        width = box[axis+2]-box[axis]
        if span.start <= center <= span.end or width + 2*INLINE_CLEARANCE > span.end-span.start:
            continue
        overhang = _span_overhang(span, strokes)
        if overhang < .003:
            continue
        original = optional(item.annotation, "GetPosition")
        if not original or len(original) != 3:
            continue
        delta = (span.start+span.end)/2-center
        others = occupied[:index] + occupied[index+1:]
        used = [label for label, _ in others]
        own = outlines[name]
        model = edges.get(name, ())
        foreign = (*reserved, *(outline for owner, outline in outlines.items() if owner != name))
        across = 1-axis
        side = -1 if span.baseline < own[across] else 1 if span.baseline > own[across+2] else 0
        attempts = (0., -.006*side, .006*side, -.012*side, .012*side) if side else (0.,)
        shifted = False
        for step in attempts:
            lane = span.baseline + step
            if side and side*(lane-own[across+(2 if side > 0 else 0)]) <= .001:
                continue
            offset = (delta, step) if axis == 0 else (step, delta)
            if not _fits_detail(_translated_box(box, *offset), allowed, foreign, used, own, model):
                continue
            target = (original[0]+offset[0], original[1]+offset[1], original[2])
            if not optional(item.annotation, "SetPosition2", *target):
                continue
            shifted = True
            actual, _, _ = text_geometry(item.annotation)
            lines = dimension_lines(item.annotation)
            updated = _linear_span(item)
            if (updated is not None and updated.axis == axis
                    and max(abs(updated.start-span.start), abs(updated.end-span.end),
                            abs(updated.baseline-lane)) <= CHAIN_TOLERANCE
                    and actual is not None and span.start + .001 <= actual[axis]
                    and actual[axis+2] <= span.end - .001
                    and _span_overhang(updated, lines) < overhang - .001
                    and _fits_detail(actual, allowed, foreign, used, own, model)
                    and _routes_fit(actual, lines, foreign, others)):
                moved += 1
                occupied[index] = (actual, lines)
                placed[index] = (item, name, updated)
                break
        else:
            if shifted and not optional(item.annotation, "SetPosition2", *original):
                raise RuntimeError("SolidWorks не восстановил положение размера")
    return moved


def _callout_crowding(box, used):
    gaps = (hypot(max(other[0]-box[2], box[0]-other[2], 0.),
                  max(other[1]-box[3], box[1]-other[3], 0.)) for other in used)
    weights = [(1-gap/CALLOUT_RADIUS)**2 for gap in gaps if gap < CALLOUT_RADIUS]
    return sum(weights), len(weights)


def _callout_positions(box, view, allowed, foreign, used, score):
    x, y = (box[0]+box[2])/2, (box[1]+box[3])/2
    half_w, half_h = (box[2]-box[0])/2, (box[3]-box[1])/2
    x0, y0, x1, y1 = view
    targets = []
    for fraction in (.12, .35, .65, .88):
        xx, yy = x0+fraction*(x1-x0), y0+fraction*(y1-y0)
        for gap in (.005, .010):
            targets.extend(((xx, y0-half_h-gap), (xx, y1+half_h+gap),
                            (x0-half_w-gap, yy), (x1+half_w+gap, yy)))
    options = []
    for xx, yy in targets:
        dx, dy = xx-x, yy-y
        candidate = _translated_box(box, dx, dy)
        if not fits(candidate, allowed, (view, *foreign), used):
            continue
        cost = _callout_crowding(candidate, used)[0] + hypot(dx, dy)*5
        if cost < score-.12:
            options.append((cost, dx, dy))
    return sorted(options)[:6]


def _line_crossings(strokes, others):
    def side(a, b, p):
        return (b[0]-a[0])*(p[1]-a[1]) - (b[1]-a[1])*(p[0]-a[0])

    return sum(side(a, b, c)*side(a, b, d) < -1e-12 and side(c, d, a)*side(c, d, b) < -1e-12
               for a, b in strokes for c, d in others)


def _single_arrow(item):
    data = optional(item.annotation, "GetDisplayData")
    if data is not None and optional(data, "GetArrowHeadCount") == 1:
        arrow = optional(data, "GetArrowHeadAtIndex2", 0)
        if arrow and len(arrow) >= 5 and all(map(isfinite, arrow[:5])):
            return arrow
    return None


def _distance_to_stroke(point, stroke):
    (x0, y0), (x1, y1) = stroke
    dx, dy = x1-x0, y1-y0
    length_sq = dx*dx + dy*dy
    if not length_sq:
        return dist(point, stroke[0])
    projection = ((point[0]-x0)*dx + (point[1]-y0)*dy) / length_sq
    fraction = min(max(projection, 0.), 1.)
    return dist(point, (x0+fraction*dx, y0+fraction*dy))


def _in_arc_span(point, center, stroke):
    # Exported polylines can approximate a bend with only a few chords.
    # Compare the arrow's angle to the endpoints rather than its distance to a chord.
    (ax, ay), (bx, by), (px, py) = ((x-center[0], y-center[1]) for x, y in (*stroke, point))
    turn = ax*by-ay*bx
    return abs(turn) > 1e-12 and turn*(ax*py-ay*px) >= 0. and turn*(px*by-py*bx) >= 0.


def _radius_arc(arrow, radius, geometry, *, allow_extension=False):
    if arrow is None or radius <= 0 or (direction := hypot(*arrow[3:5])) < .99:
        return (), 0.
    point = arrow[:2]
    dx, dy = (component*radius/direction for component in arrow[3:5])
    tolerance = max(.0001, min(.0004, radius*.02))
    # SolidWorks can put a radial arrow on the extension of the supporting
    # circle. Verify the original visible arc rather than the infinite circle.
    closest = (), float("inf")
    for center in ((point[0]+dx, point[1]+dy), (point[0]-dx, point[1]-dy)):
        arc = tuple(stroke for stroke in geometry
                    if all(abs(dist(end, center)-radius) <= tolerance for end in stroke))
        if any(_in_arc_span(point, center, stroke) for stroke in arc):
            return arc, tolerance
        gap = min((_distance_to_stroke(point, stroke) for stroke in arc), default=float("inf"))
        if gap < closest[1]:
            closest = arc, gap
    arc, gap = closest
    return (arc if allow_extension or gap <= tolerance else ()), tolerance


def _doubled_radius_leader(arrow, radius, strokes):
    if arrow is None or radius <= 0:
        return False
    x, y = arrow[:2]
    dx, dy = arrow[3:5]
    branches = []
    for stroke in strokes:
        if min(dist((x, y), p) for p in stroke) > min(.003, .25*radius):
            continue
        a, b = sorted(((px-x)*dx+(py-y)*dy, (py-y)*dx-(px-x)*dy) for px, py in stroke)
        length = b[0]-a[0]
        if length > .002 and abs(b[1]-a[1]) < .65*length:
            branches.append((a, b, (b[1]-a[1])/length))
    # The outward strokes overlap along the arrow direction but run beside
    # one another. The line to the center and the text shelf meet end to end.
    for index, (a, b, slope) in enumerate(branches):
        for c, d, other_slope in branches[index+1:]:
            low, high = max(a[0], c[0]), min(b[0], d[0])
            if high-low > .001 and any(abs((a[1]+(t-a[0])*slope) -
                                           (c[1]+(t-c[0])*other_slope)) > .00015
                                       for t in (low, high)):
                return True
    return False


def _doubled_diameter_leader(item, strokes):
    if item.kind != 6:
        return False
    data = optional(item.annotation, "GetDisplayData")
    if data is None or optional(data, "GetArrowHeadCount") != 2:
        return False
    arrows = [optional(data, "GetArrowHeadAtIndex2", i) for i in range(2)]
    if not all(arrow and len(arrow) >= 2 and all(map(isfinite, arrow[:2])) for arrow in arrows):
        return False
    tips = [arrow[:2] for arrow in arrows]
    radius = dist(*tips) / 2
    center = tuple((a+b)/2 for a, b in zip(*tips))
    tolerance = min(.001, .10*radius)
    branches = []
    for line in strokes:
        first, last = sorted(line, key=lambda p: dist(p, center))
        if abs(dist(first, center)-radius) <= tolerance and dist(last, center) > radius+.002:
            branches.append((first, last))
    # A two-arrow diameter has one outward branch on each side. A second
    # branch leaving the same arc and diverging beside the text is stray.
    return any(dist(a, c) <= .5*radius and dist(b, d) > tolerance
               for i, (a, b) in enumerate(branches) for c, d in branches[i+1:])


def _spread_callouts(placed, occupied, outlines, edges, reserved, allowed):
    moved = 0
    for index, (item, name, _) in enumerate(placed):
        if item.kind not in (5, 6):
            continue
        own = outlines[name]
        box, strokes = occupied[index]
        if not strokes:
            continue
        others = occupied[:index] + occupied[index+1:]
        used = [label for label, _ in others]
        score, count = _callout_crowding(box, used)
        if item.kind == 5:
            value = optional(item.dimension, "GetSystemValue2", "")
            scale = optional(item.view, "ScaleDecimal")
            radius = (value or 0.)*(scale or 0.)
            arrow = _single_arrow(item)
            doubled = _doubled_radius_leader(arrow, radius, strokes)
        else:
            doubled = _doubled_diameter_leader(item, strokes)
        if not doubled and (intersects(box, own) or count < 2 or score < .25):
            continue
        foreign = (*reserved, *(outline for owner, outline in outlines.items() if owner != name))
        options = _callout_positions(box, own, allowed, foreign, used,
                                     float("inf") if doubled else score)
        if doubled:
            steps = (.004, -.004, .006, -.006, .008, -.008, .014, -.014)
            nearby = [(0., dx, dy) for delta in steps
                      for dx, dy in ((0., delta), (delta, 0.))]
            if item.kind == 5:
                shifts = (-.002, .002, -.004, .004, -.008, .008, -.012, .012, -.016, .016)
                nearby.extend((0., dx, dy) for dx in shifts for dy in steps)
            options = nearby + options
        if not options:
            continue
        original = optional(item.annotation, "GetPosition")
        if not original or len(original) != 3:
            continue
        neighboring_lines = tuple(line for _, lines in others for line in lines)
        geometry = edges.get(name, ())
        if item.kind == 5:
            arc, _ = _radius_arc(arrow, radius, geometry, allow_extension=doubled)
            if not arc:
                continue
        crossings = _line_crossings(strokes, neighboring_lines)
        geometry_crossings = _line_crossings(strokes, geometry)
        length = sum(dist(*line) for line in strokes)
        shifted = False
        for _, dx, dy in options:
            if not optional(item.annotation, "SetPosition2", original[0]+dx, original[1]+dy, original[2]):
                continue
            shifted = True
            actual, _, _ = text_geometry(item.annotation)
            lines = dimension_lines(item.annotation)
            if item.kind == 5:
                arrow = _single_arrow(item)
                if (arrow is None or _doubled_radius_leader(arrow, radius, lines)
                        or not _radius_arc(arrow, radius, arc)[0]):
                    continue
            if (fits(actual, allowed, (own, *foreign), used) and lines
                    and (doubled or _callout_crowding(actual, used)[0] < score-.12)
                    and sum(dist(*line) for line in lines) <= max(length*1.35, length+.018)
                    and _routes_fit(actual, lines, foreign, others,
                                    text_clearance=.002 if doubled else 0.)
                    and _line_crossings(lines, neighboring_lines) <= crossings
                    and _line_crossings(lines, geometry) <= geometry_crossings
                    and not _doubled_diameter_leader(item, lines)
                    and not optional(item.annotation, "IsDangling")):
                occupied[index] = (actual, lines)
                moved += 1
                break
        else:
            if shifted and not optional(item.annotation, "SetPosition2", *original):
                raise RuntimeError("SolidWorks не восстановил положение выноски")
    return moved


def arrange_dimensions(doc, frame):
    views = {get_value(view, "GetName2"): view for view in drawing_views(doc)}
    boxes = {name: get_value(view, "GetOutline")[:4] for name, view in views.items()}
    get_value(doc, "ViewZoomtofit2")
    view_edges = {name: view_strokes(view, boxes[name]) for name, view in views.items()}
    flat_views = {name for name, view in views.items() if optional(view, "IsFlatPatternView")}
    model_edges = {name: view_edges[name]
                   for name, view in views.items() if get_value(view, "Type") != 2
                   and name not in flat_views}
    outlines = {name: geometry_outline(view_edges[name], box) for name, box in boxes.items()}
    reserved = list(format_zones(doc, frame))
    reserved.extend(zone for view in views.values() for zone in section_line_zones(view))
    for name, view in views.items():
        for caption in section_captions(view):
            box = _place_caption(caption, outlines[name], frame.drawable, outlines.values(), reserved)
            if box is not None:
                reserved.append(box)
        if name in flat_views:
            for raw in optional(view, "GetNotes") or ():
                note = as_dispatch(raw)
                if optional(note, "GetText") == FLAT_PATTERN_CAPTION:
                    annotation = optional(note, "GetAnnotation")
                    box = _place_flat_caption(annotation, outlines[name]) if annotation is not None else None
                    if box is not None:
                        reserved.append(box)
    occupied = []
    accepted = []
    seen = set()
    hidden = moved = 0
    rows = []
    for item in visible_dimensions(doc):
        full_name = optional(item.dimension, "FullName") if item.dimension is not None else ""
        # Drawing dimensions are placed using actual geometry. Follow them
        # with explicitly marked model dimensions in size/kind order.
        driven = str(full_name or "").lower().endswith(".drawing")
        display_data = optional(item.annotation, "GetDisplayData")
        content, angle, text = text_geometry(item.annotation, display_data=display_data)
        family = optional(item.dimension, "GetSystemValue2", "") if item.dimension is not None else None
        kind = 2 if item.kind in (11, 12) else item.kind
        own_name = get_value(item.view, "GetName2")
        vertical = kind == 2 and (item.kind == 12 or (angle is not None and abs(cos(angle)) < .7))
        key = (own_name, kind, round(family, 6), vertical, text) if isinstance(family, (int, float)) else None
        span = _linear_span(item, display_data=display_data)
        if key is not None and span is not None:
            key += (span.axis, round(span.start/CHAIN_TOLERANCE), round(span.end/CHAIN_TOLERANCE))
        rows.append((own_name in flat_views, not driven, item.kind not in (5, 6),
                     item, content, angle, key, own_name, span))
    rows.sort(key=lambda row: row[:3])
    for _, _, _, item, box, angle, key, own_name, span in rows:
        if key in seen:
            _hide(item.annotation)
            hidden += 1
            continue
        used = [row[0] for row in occupied]
        view = boxes[own_name]
        foreign = (*reserved, *(outline for name, outline in boxes.items() if name != own_name))
        edges = model_edges.get(own_name, ())
        strokes = dimension_lines(item.annotation)
        if (not _fits_detail(box, frame.drawable, foreign, used, view, edges)
                or not _routes_fit(box, strokes, foreign, occupied)):
            anchor = optional(item.annotation, "GetPosition")
            if not anchor or len(anchor) < 3 or box is None:
                _hide(item.annotation)
                hidden += 1
                continue
            max_length = max((dist(*line) for line in strokes), default=0.)
            outer = candidates(anchor[:3], box, view, item.kind, angle)
            inner = _interior_candidates(anchor[:3], box, view) if edges else ()
            options = (*inner, *outer) if intersects(box, view) else (*outer, *inner)
            for position, predicted in options:
                if (not _fits_detail(predicted, frame.drawable, foreign, used, view, edges)
                        or not _routes_fit(predicted, (), (), occupied)):
                    continue
                if optional(item.annotation, "SetPosition2", *position) is False:
                    continue
                placed, _, _ = text_geometry(item.annotation)
                if not _fits_detail(placed, frame.drawable, foreign, used, view, edges):
                    continue
                moved_lines = dimension_lines(item.annotation)
                longest = max((dist(*line) for line in moved_lines), default=0.)
                if (longest <= max_length + .045
                        and _routes_fit(placed, moved_lines, foreign, occupied)):
                    box = placed
                    strokes = moved_lines
                    if span is not None:
                        span = _linear_span(item)
                    moved += 1
                    break
            else:
                _hide(item.annotation)
                hidden += 1
                continue
        occupied.append((box, strokes))
        accepted.append((item, own_name, span))
        if key is not None:
            seen.add(key)
    # Flat patterns keep the two independently positioned blank extents.
    chain_views = {name: outline for name, outline in outlines.items() if name not in flat_views}
    moved += _center_linear_labels(accepted, occupied, outlines, model_edges, reserved, frame.drawable)
    moved += _align_chains(accepted, occupied, chain_views, (*reserved, *(outlines[n] for n in flat_views)),
                           frame.drawable)
    moved += _spread_callouts(accepted, occupied, outlines, view_edges, reserved, frame.drawable)
    return len(occupied), moved, hidden


class DimensionTextLayout:
    """Place a longer label while keeping the existing dimensions in place."""

    def __init__(self, doc, frame):
        self.frame = frame.drawable
        views = {get_value(view, "GetName2"): view for view in drawing_views(doc)}
        self.boxes = {name: get_value(view, "GetOutline")[:4] for name, view in views.items()}
        self.edges = {name: view_strokes(view, self.boxes[name]) for name, view in views.items()
                      if get_value(view, "Type") != 2 and not optional(view, "IsFlatPatternView")}
        self.reserved = list(format_zones(doc, frame))
        self.reserved.extend(zone for view in views.values() for zone in section_line_zones(view))
        for view in views.values():
            for raw in optional(view, "GetNotes") or ():
                note = as_dispatch(raw)
                annotation = optional(note, "GetAnnotation")
                if annotation is not None and (box := text_geometry(annotation)[0]) is not None:
                    self.reserved.append(box)
        self.occupied = {}
        self.items = tuple(visible_dimensions(doc))
        for item in self.items:
            if (box := text_geometry(item.annotation)[0]) is not None:
                self.occupied[self._key(item.view, item.annotation)] = box, dimension_lines(item.annotation)

    @staticmethod
    def _key(view, annotation):
        display = optional(annotation, "GetSpecificAnnotation")
        dimension = optional(display, "GetDimension2", 0) if display is not None else None
        # Imported dimensions in different model features can both be named D2.
        full_name = optional(dimension, "FullName") if dimension is not None else None
        return get_value(view, "GetName2"), full_name or get_value(annotation, "GetName")

    def place(self, view, annotation, kind):
        key = self._key(view, annotation)
        name = key[0]
        used = [row for entry, row in self.occupied.items() if entry != key]
        foreign = (*self.reserved, *(box for other, box in self.boxes.items() if other != name))
        own = self.boxes[name]
        edges = self.edges.get(name, ())
        origin = optional(annotation, "GetPosition")
        box, angle, _ = text_geometry(annotation)
        if not origin or box is None:
            return False

        def accept():
            actual, _, _ = text_geometry(annotation)
            strokes = dimension_lines(annotation)
            if (_fits_detail(actual, self.frame, foreign, [box for box, _ in used], own, edges)
                    and _routes_fit(actual, strokes, foreign, used)):
                self.occupied[key] = actual, strokes
                return True
            return False

        if accept():
            return True
        outer = candidates(origin, box, own, kind, angle)
        inner = _interior_candidates(origin, box, own) if edges else ()
        for position, predicted in (*outer, *inner):
            if not _fits_detail(predicted, self.frame, foreign, [box for box, _ in used], own, edges):
                continue
            if optional(annotation, "SetPosition2", *position) and accept():
                return True
        optional(annotation, "SetPosition2", *origin)
        return False

    def free_diameter_routes(self, view, original, anchors, label_size, centerlines, *,
                             used_centers=(), allow_outline_exit=False, refine=False):
        """Suggest short routes, then check the geometry actually drawn by SolidWorks."""
        name = get_value(view, "GetName2")
        key = self._key(view, original) if original is not None else None
        edges = self.edges.get(name, ())
        if not edges:  # No reliable way to check the leader against this view.
            return
        occupied = [row for entry, row in self.occupied.items() if entry != key]
        boxes = [box for box, _ in occupied]
        foreign = (*self.reserved, *(box for other, box in self.boxes.items() if other != name))
        obstacles = tuple(chain(edges, centerlines, *(lines for _, lines in occupied)))
        own = self.boxes[name]
        border = ()
        if allow_outline_exit:
            own = geometry_outline(edges, own)
            # Permit one exit through a straight exterior edge. Inner edges,
            # centerlines, and other annotations are still obstacles.
            border = tuple(edge for edge in edges if any(
                all(abs(p[axis]-own[bound]) <= .0001 for p in edge)
                for bound, axis in enumerate((0, 1, 0, 1))))
        def accept(annotation, center, radius):
            box = text_geometry(annotation)[0]
            strokes = dimension_lines(annotation)
            return (fits(box, self.frame, foreign, boxes)
                    and not any(crosses(edge, box) for edge in obstacles)
                    and _routes_fit(box, strokes, foreign, occupied)
                    and _clear_leader(strokes, obstacles, center, radius+.0004,
                                      border=border if not intersects(box, own) else ()))

        for fine in ((False, True) if refine else (False,)):
            proposals = _leader_positions(anchors, *label_size, self.frame, own,
                                          foreign, boxes, obstacles, border=border, refine=fine)
            if used_centers:
                free = (route for route in proposals if route[4] not in used_centers)
                used = (route for route in proposals if route[4] in used_centers)
                candidates = chain(islice(free, 24), islice(used, 24))
            else:
                candidates = islice(proposals, 24)
            for _, _, anchor, position, center, radius in candidates:
                yield anchor, position, center, radius, accept
