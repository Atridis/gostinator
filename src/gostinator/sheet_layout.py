"""Choose a native GOST sheet format, readable scale, and view positions."""
from dataclasses import dataclass
from fractions import Fraction
import os
from pathlib import Path

from pythoncom import com_error

from .model_geometry import choose_section, needs_reverse, primary_axis
from .solidworks import get_value


@dataclass(frozen=True)
class Paper:
    name: str
    width: float
    height: float
    sw_code: int

    @property
    def drawable(self):
        # ESKD sheet margins: 20 mm on the binding side and 5 mm elsewhere.
        # A first-sheet title block occupies the bottom 55 mm of the frame.
        return .020, .060, self.width - .005, self.height - .005


PAPERS = (
    Paper("A4", .210, .297, 7),
    Paper("A3", .420, .297, 8),
    Paper("A2", .594, .420, 9),
    Paper("A1", .841, .594, 10),
    Paper("A0", 1.189, .841, 11),
)
SCALES = (1., .5, .4, .25, .2, .1, .05, .025)
VIEW_AXES = ((2, 1), (0, 2), (0, 1))
EDGE_LANE = .014


@dataclass(frozen=True)
class ViewSlot:
    axis: int
    sign: int
    column: int
    row: int
    role: str = "ordinary"
    turn: bool = False
    factor: float = 1.
    align_right: bool = False


@dataclass(frozen=True)
class Layout:
    paper: Paper
    scale: float
    positions: tuple[tuple[float, float], ...]
    slots: tuple[ViewSlot, ...]


def view_slots(model, first_angle, flat):
    axis = primary_axis(model)
    cut = choose_section(model, axis)
    if cut:
        normal = next(i for i in range(3) if i not in (axis, cut.along))
        horizontal = cut.along == VIEW_AXES[axis][0]
        central_through_bore = (len({b.center for b in cut.bores}) == 1
                                and any(b.high-b.low >= .85*model.spans[axis] for b in cut.bores))
        turn = not horizontal and central_through_bore
        if horizontal or turn:
            top = int(first_angle)
            return (ViewSlot(axis, 1, 0, top, turn=turn),
                    ViewSlot(normal, 1, 0, 1-top, "section", turn=turn))
        left = int(not first_angle)
        return (ViewSlot(axis, 1, left, 0), ViewSlot(normal, 1, 1-left, 0, "section"))
    axes = sorted((i for i in range(3) if i != axis), key=lambda i: model.spans[i])
    side, other = axes
    side_sign = -1 if side == 0 else 1
    if model.sheet_metal:
        offset = int(flat)
        folded = (ViewSlot(side, side_sign, 0, 1+offset), ViewSlot(axis, 1, 0, offset))
        if flat:
            return (*folded, ViewSlot(axis, 1, 0, 0, "flat", factor=.5, align_right=True))
        return folded
    if needs_reverse(model, axis):
        return (ViewSlot(axis, 1, 0, 0), ViewSlot(side, side_sign, 1, 0),
                ViewSlot(axis, -1, 2, 0))
    return (ViewSlot(axis, 1, 0, 1), ViewSlot(side, side_sign, 1, 1),
            ViewSlot(other, -1 if other == 0 else 1, 0, 0))


def _allocate(limit_low, limit_high, spans, gap):
    available = limit_high - limit_low - 2 * EDGE_LANE
    total = sum(spans) + gap * (len(spans) - 1)
    if total > available + 1e-9:
        return None
    # Use part of the remaining room to separate the labels on adjacent views.
    surplus = available - total
    if len(spans) > 1:
        gap += min(.035, surplus / len(spans))
        total = sum(spans) + gap * (len(spans)-1)
    cursor = (limit_low + limit_high - total) / 2
    offsets = []
    for span in spans:
        offsets.append(cursor+span/2)
        cursor += span+gap
    return offsets


def pack(paper, scale, slots, extents, gap):
    widths = [0.] * (max(slot.column for slot in slots) + 1)
    heights = [0.] * (max(slot.row for slot in slots) + 1)
    for slot, (w, h) in zip(slots, extents):
        widths[slot.column] = max(widths[slot.column], scale * slot.factor * w)
        heights[slot.row] = max(heights[slot.row], scale * slot.factor * h)
    left, bottom, right, top = paper.drawable
    xs = _allocate(left, right, widths, gap)
    ys = _allocate(bottom, top, heights, gap)
    if xs is not None and ys is not None:
        return tuple((right - EDGE_LANE - w*scale*s.factor/2 if s.align_right else xs[s.column],
                      ys[s.row]) for s, (w, _) in zip(slots, extents))
    return None


def choose_layout(model, first_angle, flat, *, papers=PAPERS):
    slots = view_slots(model, first_angle, flat)
    extents = []
    for slot in slots:
        axes = VIEW_AXES[slot.axis]
        extents.append(tuple(model.spans[coordinate] for coordinate in (axes[::-1] if slot.turn else axes)))
    if flat:
        # The flat blank can grow in the bend direction; check its actual
        # outline after inserting it and before placing any dimensions.
        extents[-1] = (extents[-1][0] * 1.15, extents[-1][1] * 1.15)
    gap = .024 + min(.025, len(model.bores) * .00055)
    choices = [(paper, scale) for paper in papers
               for scale in ((4., 2., 1.) if paper.name == "A4" else (1.,))]
    choices.extend((paper, scale) for scale in SCALES[1:] for paper in papers)
    for paper, scale in choices:
        positions = pack(paper, scale, slots, extents, gap)
        if positions is not None:
            return Layout(paper, scale, positions, slots)
    raise RuntimeError("Модель и размерные поля не помещаются на доступные листы")


def _format_paths(app):
    locations = []
    try:
        configured = get_value(app, "GetUserPreferenceStringValue", 7)  # swFileLocationsSheetFormat
    except (AttributeError, com_error):
        configured = ""
    for raw in (configured or "").split(";"):
        if raw.strip():
            locations.append(Path(os.path.expandvars(raw.strip().strip('"'))))
    program_data = os.environ.get("PROGRAMDATA")
    if program_data:
        root = Path(program_data) / "SOLIDWORKS"
        try:
            major = int(str(get_value(app, "RevisionNumber")).split(".")[0])
            current = root / f"SOLIDWORKS {major + 1992}"
        except (AttributeError, TypeError, ValueError, com_error):
            current = None
        releases = sorted(root.glob("SOLIDWORKS 20[0-9][0-9]"), reverse=True)
        if current is not None:
            releases = [current, *(p for p in releases if p != current)]
        for release in releases:
            for language in ("russian", "english"):
                locations.append(release / "lang" / language / "sheetformat")
    return locations


def installed_formats(app):
    folders = _format_paths(app)
    result = {}
    for paper in PAPERS:
        filename = f"{paper.name.lower()} - gost_sh1.slddrt"
        for folder in folders:
            path = folder / filename
            if path.is_file():
                result[paper] = path
                break
    return result


def _ratio(scale):
    ratio = Fraction(scale).limit_denominator(1000)
    return float(ratio.numerator), float(ratio.denominator)


def configure_sheet(doc, layout, format_file, first_angle):
    sheet = get_value(doc, "GetCurrentSheet")
    scale1, scale2 = _ratio(layout.scale)
    paper = layout.paper
    original = get_value(sheet, "GetProperties2")
    if max(abs(original[5]-paper.width), abs(original[6]-paper.height)) <= .0005:
        sketch = get_value(sheet, "GetTemplateSketch")
        if sketch is not None and get_value(sketch, "GetSketchSegments"):
            sheet.SheetFormatVisible = True
            if get_value(sheet, "SetScale", scale1, scale2, False, False) is False:
                current = get_value(sheet, "GetProperties2")
                if abs(current[2]/current[3] - layout.scale) > 1e-6:
                    raise RuntimeError("SolidWorks не применил масштаб существующего листа")
            return
    if not get_value(doc, "SetupSheet5", get_value(sheet, "GetName"), paper.sw_code, 12,
                     scale1, scale2, first_angle, str(format_file), paper.width, paper.height,
                     "", False):
        raise RuntimeError(f"SolidWorks не применил формат листа {paper.name}")
    sheet = get_value(doc, "GetCurrentSheet")
    sheet.SheetFormatVisible = True
    get_value(doc, "EditRebuild3")
    props = get_value(sheet, "GetProperties2")
    if max(abs(props[5]-paper.width), abs(props[6]-paper.height)) > .0005:
        raise RuntimeError(f"Размер листа SolidWorks отличается от формата {paper.name}")
    sketch = get_value(sheet, "GetTemplateSketch")
    if sketch is None or not get_value(sketch, "GetSketchSegments"):
        raise RuntimeError(f"SolidWorks не загрузил рамку {paper.name} из {format_file.name}")


def synchronize_sheet_scale(doc, scale):
    # Inserting the first model view can auto-scale the sheet. The individual
    # views already have explicit scales, which SetScale must leave alone.
    sheet = get_value(doc, "GetCurrentSheet")
    get_value(sheet, "SetScale", *_ratio(scale), False, False)
    get_value(doc, "EditRebuild3")
    props = get_value(sheet, "GetProperties2")
    if abs(props[2] / props[3] - scale) > 1e-6:
        raise RuntimeError("Масштаб в основной надписи не совпадает с масштабом главного вида")
