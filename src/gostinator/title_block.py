"""Fit linked notes within the two text cells of a GOST form 1 title block."""
from math import isfinite
import re

from pythoncom import com_error

from .annotation_layout import optional, text_geometry
from .solidworks import as_dispatch, get_value


SHEET_LINE = 0
EDGE_TOLERANCE = .00025
PADDING = .0015
LINE_GAP = .001
PROPERTY = re.compile(r'\$PRP(?:SHEET|VIEW)?:"([^"]+)"', re.IGNORECASE)
PROPERTY_CELLS = {
    "наименование": "name", "description": "name", "вид_документа": "name",
    "материал": "material", "material": "material", "sw-material": "material",
    "заготовка": "material", "типоразмер": "material",
}


def _sheet_lines(sheet):
    sketch = optional(sheet, "GetTemplateSketch")
    if sketch is None:
        return
    for raw in optional(sketch, "GetSketchSegments") or ():
        line = as_dispatch(raw)
        if optional(line, "GetType") != SHEET_LINE:
            continue
        ends = [optional(line, method) for method in ("GetStartPoint2", "GetEndPoint2")]
        if all(end is not None for end in ends):
            points = tuple(tuple(optional(end, axis) for axis in ("X", "Y")) for end in ends)
            if all(isinstance(v, (float, int)) and isfinite(v) for point in points for v in point):
                yield points


def form_cells(width, lines):
    # The text column of form 1 is 70 mm wide, with a 50 mm column to
    # its right. Verify the sheet-format lines before editing any notes.
    left, right = width - .125, width - .055
    cells = {"name": (left, .020, right, .045), "material": (left, .005, right, .020)}

    def has_edge(axis, coordinate, low, high):
        return any(abs(a[axis] - coordinate) <= EDGE_TOLERANCE
                   and abs(b[axis] - coordinate) <= EDGE_TOLERANCE
                   and min(a[1-axis], b[1-axis]) <= low + EDGE_TOLERANCE
                   and max(a[1-axis], b[1-axis]) >= high - EDGE_TOLERANCE for a, b in lines)

    return {name: box for name, box in cells.items()
            if all(has_edge(0, x, box[1], box[3]) for x in (box[0], box[2]))
            and all(has_edge(1, y, box[0], box[2]) for y in (box[1], box[3]))}


def _linked_notes(drawing, cells):
    groups = {name: [] for name in cells}
    sheet_view = get_value(drawing, "GetFirstView")
    for raw in optional(sheet_view, "GetNotes") or ():
        note = as_dispatch(raw)
        key = PROPERTY.fullmatch(str(optional(note, "PropertyLinkedText") or "").strip())
        cell = PROPERTY_CELLS.get(key[1].casefold()) if key else None
        if cell not in cells or not str(optional(note, "GetText") or "").strip():
            continue
        ann = optional(note, "GetAnnotation")
        position = optional(ann, "GetPosition") if ann is not None else None
        left, bottom, right, top = cells[cell]
        if position and left <= position[0] <= right and bottom <= position[1] <= top:
            groups[cell].append(ann)
    return groups


def _measure(notes):
    return [(note, box) for note in notes if (box := text_geometry(note)[0]) is not None]


def _fit_cell(drawing, notes, cell):
    measured = _measure(notes)
    if len(measured) != len(notes):
        return False
    width = cell[2] - cell[0] - 2 * PADDING
    height = cell[3] - cell[1] - 2 * PADDING
    gap = LINE_GAP * (len(measured) - 1)
    factor = min(1., width / max(b[2] - b[0] for _, b in measured),
                 (height - gap) / sum(b[3] - b[1] for _, b in measured))
    if factor <= 0:
        return False
    if factor < 1.:
        for note, _ in measured:
            fmt = get_value(note, "GetTextFormat", 0)
            fmt.CharHeight = get_value(fmt, "CharHeight") * factor * .99
            if not get_value(note, "SetTextFormat", 0, False, fmt):
                return False
        get_value(drawing, "GraphicsRedraw2")
        measured = _measure(notes)
        if len(measured) != len(notes):
            return False
    measured.sort(key=lambda item: item[1][3], reverse=True)
    total_height = sum(b[3] - b[1] for _, b in measured) + gap
    top = (cell[1] + cell[3] + total_height) / 2
    for note, b in measured:
        x, y, z = get_value(note, "GetPosition")
        dx = (cell[0] + cell[2] - b[0] - b[2]) / 2
        dy = top - b[3]
        if max(abs(dx), abs(dy)) > .0001 and not get_value(note, "SetPosition2", x+dx, y+dy, z):
            return False
        top -= b[3] - b[1] + LINE_GAP
    get_value(drawing, "GraphicsRedraw2")
    fitted = _measure(notes)
    return len(fitted) == len(notes) and all(
        cell[0] <= b[0] and cell[1] <= b[1] and b[2] <= cell[2] and b[3] <= cell[3]
        for _, b in fitted)


def fit_title_block(drawing, frame):
    sheet = get_value(drawing, "GetCurrentSheet")
    cells = form_cells(frame.width, tuple(_sheet_lines(sheet)))
    if not cells:
        return
    groups = _linked_notes(drawing, cells)
    if not any(groups.values()):
        return
    get_value(drawing, "EditTemplate")
    try:
        for name, notes in groups.items():
            try:
                if notes and not _fit_cell(drawing, notes, cells[name]):
                    print("[WARN] Надпись не поместилась в ячейку основной надписи", flush=True)
            except (AttributeError, com_error) as exc:
                print(f"[WARN] Не удалось выровнять основную надпись — {exc}", flush=True)
    finally:
        get_value(drawing, "EditSheet")

