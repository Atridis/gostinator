"""Arrange dimension labels against native text and view geometry."""
from collections import namedtuple
from math import cos, dist, isfinite, sin
from pathlib import Path

from pythoncom import VT_BYREF, VT_VARIANT, com_error
from win32com.client import VARIANT

from .drawing_views import drawing_views
from .solidworks import as_dispatch, get_value


VisibleDimension = namedtuple("VisibleDimension", "view display annotation dimension kind")


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


def text_geometry(annotation):
    data = optional(annotation, "GetDisplayData")
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


def _routes_fit(box, strokes, foreign_views, occupied):
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


def section_captions(view):
    if get_value(view, "Type") != 2:
        return
    for raw in optional(view, "GetNotes") or ():
        note = as_dispatch(raw)
        linked = str(optional(note, "PropertyLinkedText") or "")
        if "<VLLABEL>" in linked:
            annotation = optional(note, "GetAnnotation")
            if annotation is not None and optional(annotation, "Visible") == 1:
                yield annotation


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
    for name, view in views.items():
        for caption in section_captions(view):
            box = _place_caption(caption, outlines[name], frame.drawable, outlines.values(), reserved)
            if box is not None:
                reserved.append(box)
    occupied = []
    seen = set()
    hidden = moved = 0
    rows = []
    for item in visible_dimensions(doc):
        full_name = optional(item.dimension, "FullName") if item.dimension is not None else ""
        # Drawing dimensions are placed using actual geometry. Follow them
        # with explicitly marked model dimensions in size/kind order.
        driven = str(full_name or "").lower().endswith(".drawing")
        content, angle, text = text_geometry(item.annotation)
        family = optional(item.dimension, "GetSystemValue2", "") if item.dimension is not None else None
        kind = 2 if item.kind in (11, 12) else item.kind
        own_name = get_value(item.view, "GetName2")
        vertical = kind == 2 and (item.kind == 12 or (angle is not None and abs(cos(angle)) < .7))
        key = (own_name, kind, round(family, 6), vertical, text) if isinstance(family, (int, float)) else None
        rows.append((own_name in flat_views, not driven, item.kind not in (5, 6),
                     item, content, angle, key, own_name))
    rows.sort(key=lambda row: row[:3])
    for _, _, _, item, box, angle, key, own_name in rows:
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
                    moved += 1
                    break
            else:
                _hide(item.annotation)
                hidden += 1
                continue
        occupied.append((box, strokes))
        if key is not None:
            seen.add(key)
    return len(occupied), moved, hidden
