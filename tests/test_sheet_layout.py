from math import cos, pi, sin
from pathlib import Path
from types import SimpleNamespace

import pytest

from gostinator import annotation_layout
from gostinator.annotation_layout import (
    _align_chains, _callout_crowding, _callout_positions, _center_linear_labels, _distance_to_stroke,
    _fits_detail, _linear_span, _polyline_strokes, _radius_arc, _routes_fit, _span_overhang, _try_chain,
    candidates, crosses, dimension_lines, fits, format_zones, section_line_zones, text_extent, text_geometry,
)
from gostinator.drawing_views import Frame
from gostinator.title_block import _linked_notes, form_cells
from gostinator.model_geometry import Bore, Model
from gostinator.sheet_layout import Layout, PAPERS, choose_layout, configure_sheet


def test_large_opposed_views_use_a_larger_sheet_before_shrinking():
    model = Model((-.110, -.115, 0, .110, .115, .019), (), False,
                  ((0, 0), (0, 0), (8, 19)))
    layout = choose_layout(model, True, False)
    assert layout.paper.name == "A2"
    assert layout.scale == 1
    assert [s.sign for s in layout.slots] == [1, -1, -1]
    assert len({round(y, 6) for _, y in layout.positions}) == 1


def test_small_part_can_use_enlargement_on_existing_a4():
    model = Model((-.012, -.009, 0, .012, .009, .010),
                  (Bore(2, (-.004, 0.), .003, 0, .006),
                   Bore(2, (-.004, 0.), .005, .006, .010)), False)
    layout = choose_layout(model, True, False)
    assert layout.paper.name == "A4"
    assert layout.scale >= 2


def test_flat_pattern_uses_a_separate_scale_and_row():
    model = Model((0., 0., 0., .112, .043, .048), (), True)
    layout = choose_layout(model, True, True)
    assert layout.paper.name == "A4"
    assert layout.slots[-1].role == "flat" and layout.slots[-1].factor == .5
    assert len({y for _, y in layout.positions}) == 3


def test_sheet_metal_uses_two_folded_views_if_blank_is_unavailable():
    model = Model((0., 0., 0., .112, .043, .048), (), True)
    layout = choose_layout(model, True, False)
    upper, lower = layout.slots
    assert upper.role == lower.role == "ordinary"
    assert upper.axis != lower.axis and upper.column == lower.column
    assert layout.positions[0][1] > layout.positions[1][1]


def test_native_text_extent_respects_rotation_and_reserves_space():
    box = text_extent([((.10, .15, 0), .010, .0035, pi/2)])
    assert box == pytest.approx((.0965, .15, .1005, .16))
    assert fits(box, PAPERS[0].drawable, (), ())
    assert not fits(box, PAPERS[0].drawable, (), ((.1, .155, .11, .168),))
    assert not fits(box, PAPERS[0].drawable, ((.09, .14, .12, .18),), ())
    assert not fits((.11, .055, .13, .058), PAPERS[0].drawable, (), ())


def test_placement_can_move_a_linear_label_off_a_view():
    box = (.090, .150, .101, .154)
    view = (.050, .100, .150, .200)
    valid = [pos for pos, moved_box in candidates((.096, .15, 0), box, view, 2, 0.)
             if fits(moved_box, PAPERS[0].drawable, (view,), ())]
    assert valid
    assert all(pos[1] < view[1] or pos[1] > view[3] for pos in valid)


def test_existing_native_format_is_preserved_on_matching_sheet_size():
    scale = []
    sheet = SimpleNamespace(
        GetProperties2=(7, 12, 1., 1., True, .210, .297),
        GetTemplateSketch=SimpleNamespace(GetSketchSegments=[object()]),
        SheetFormatVisible=False, SetScale=lambda *value: scale.append(value),
    )
    drawing = SimpleNamespace(GetCurrentSheet=sheet)
    configure_sheet(drawing, Layout(PAPERS[0], 2., (), ()), Path("stock.slddrt"), True)
    assert sheet.SheetFormatVisible
    assert scale == [(2., 1., False, False)]


def test_dimension_routes_avoid_other_text_and_neighboring_views():
    label = (.02, .02, .04, .03)
    crossing = ((0., .025), (.05, .025))
    outside = ((0., .015), (.05, .015))
    neighbor = (.05, .01, .07, .03)
    assert crosses(crossing, label)
    assert not crosses(outside, label)
    assert not _routes_fit(label, (), (), [((.05, .03, .06, .04), (crossing,))])
    assert not _routes_fit(label, (((.04, .015), (.08, .015)),), [neighbor], [])
    assert _routes_fit(label, (outside,), [neighbor], [])


def linear_annotation(start, end, lane, axis=0):
    def xy(along, across):
        return (along, across, 0.) if axis == 0 else (across, along, 0.)

    position = xy((start+end)/2-.0025, lane)

    def move(*target):
        nonlocal position
        position = target
        return True

    display = SimpleNamespace(
        GetArrowHeadCount=2,
        GetArrowHeadAtIndex2=lambda index: (*xy((start, end)[index], position[1-axis]),
                                            *xy(1. if index == 0 else -1., 0.)),
        GetTextCount=1, GetTextAtIndex=lambda _: "10", GetTextInBoxWidthAtIndex=lambda _: .005,
        GetTextAngleAtIndex=lambda _: axis*pi/2, GetTextPositionAtIndex=lambda _: position,
        GetTextHeightAtIndex=lambda _: .0035, GetLineCount=1,
        GetLineAtIndex=lambda _: (0., *xy(start, position[1-axis]), *xy(end, position[1-axis])),
    )
    return SimpleNamespace(kind=2, annotation=SimpleNamespace(
        GetDisplayData=display, GetPosition=lambda: position, SetPosition2=move))


@pytest.mark.parametrize("axis", (0, 1))
def test_adjacent_linear_dimensions_share_a_row_without_changing_their_spans(axis):
    items = [linear_annotation(start, end, lane, axis) for start, end, lane in (
        (.07, .08, .075), (.08, .092, .085), (.092, .102, .066),
        (.11, .12, .075),  # Same direction and nominal, with no common witness line.
        (.07, .102, .060),  # Overall dimension spans the chain.
    )]
    occupied = [(text_geometry(item.annotation)[0], dimension_lines(item.annotation)) for item in items]
    before = [_linear_span(item) for item in items]
    assert _align_chains([(item, "view", span) for item, span in zip(items, before)], occupied,
                         {"view": (.1, .1, .2, .2)}, (), (.02, .02, .3, .3)) == 2
    after = [_linear_span(item) for item in items]
    assert [span.baseline for span in after] == pytest.approx([.085, .085, .085, .075, .060])
    assert [(s.axis, s.start, s.end) for s in after] == [(s.axis, s.start, s.end) for s in before]


def test_chain_rolls_back_if_a_label_hits_another_dimension():
    items = [linear_annotation(.06, .08, .075), linear_annotation(.08, .10, .085)]
    before = [_linear_span(item) for item in items]
    occupied = [((.086, .084, .094, .090), ())]
    assert _try_chain(list(zip(items, before)), .085, (.02, .02, .3, .3),
                      (.10, .1, .2, .2), (), occupied) is None
    assert [_linear_span(item) for item in items] == before


@pytest.mark.parametrize("axis,side", ((0, -1), (1, 1)))
def test_enclosing_dimension_moves_outside_the_smaller_chain(axis, side):
    near, inverted = ((.105, .109) if side == -1 else (.195, .191))
    items = [linear_annotation(start, end, lane, axis) for start, end, lane in (
        (.125, .157, near), (.157, .169, near), (.125, .169, inverted),
    )]
    before = [_linear_span(item) for item in items]
    occupied = [(text_geometry(item.annotation)[0], dimension_lines(item.annotation)) for item in items]
    assert _align_chains([(item, "view", span) for item, span in zip(items, before)], occupied,
                         {"view": (.12, .12, .18, .18)}, (), (.02, .02, .3, .3)) == 1
    after = [_linear_span(item) for item in items]
    assert after[:2] == before[:2]
    assert after[2].baseline == pytest.approx(near + side*.006)
    assert (after[2].start, after[2].end) == (before[2].start, before[2].end)


def test_chain_placement_follows_ordinary_label_relocation(monkeypatch):
    view = SimpleNamespace(GetName2="main", GetOutline=(.06, .12, .15, .19), Type=7)
    item = linear_annotation(.08, .10, .15)
    item.view, item.dimension = view, None
    document = SimpleNamespace(ViewZoomtofit2=lambda: None, GetCurrentSheet=SimpleNamespace())
    monkeypatch.setattr(annotation_layout, "drawing_views", lambda _: [view])
    monkeypatch.setattr(annotation_layout, "visible_dimensions", lambda _: [item])
    assert annotation_layout.arrange_dimensions(document, Frame(.210, .297, True)) == (1, 1, 0)


@pytest.mark.parametrize("length,moved", ((.020, 1), (.007, 0)))
def test_linear_text_returns_inside_only_when_it_fits(length, moved):
    item = linear_annotation(.06, .06+length, .085)
    ann = item.annotation
    ann.SetPosition2(.12, .085, 0.)
    data = ann.GetDisplayData
    first = data.GetLineAtIndex(0)
    data.GetLineCount = 2
    data.GetLineAtIndex = lambda index: first if index == 0 else (
        0., .06+length, .085, 0., *ann.GetPosition())
    span = _linear_span(item)
    # An in-span endpoint of the second segment leaves no external tail.
    placed = [(item, 'view', span)]
    occupied = [(text_geometry(ann)[0], dimension_lines(ann))]
    assert _center_linear_labels(placed, occupied, {'view': (.04, .10, .15, .18)}, {}, (),
                                 (.02, .02, .20, .25)) == moved
    if moved:
        label, lines = occupied[0]
        assert span.start < label[0] < label[2] < span.end
        assert _span_overhang(span, lines) == 0.
    else:
        assert ann.GetPosition()[0] == .12


def test_callouts_try_quieter_areas_around_the_view_with_reserved_space():
    view = (.055, .09, .155, .17)
    box = (.163, .095, .178, .100)
    used = ((.162, .1015, .176, .106), (.163, .108, .173, .112))
    reserved = ((.03, .17, .19, .23),)
    allowed = (.02, .05, .195, .23)
    score, nearby = _callout_crowding(box, used)
    assert nearby == 2
    options = _callout_positions(box, view, allowed, reserved, used, score)
    assert options
    for _, dx, dy in options:
        translated = (box[0]+dx, box[1]+dy, box[2]+dx, box[3]+dy)
        assert fits(translated, allowed, (view, *reserved), used)
        assert _callout_crowding(translated, used)[0] < score


def test_radius_move_keeps_the_arrow_on_the_visible_arc():
    points = [(cos(i*pi/8)*.010, sin(i*pi/8)*.010) for i in range(5)]
    arc = tuple(zip(points, points[1:]))
    tangent = (points[0], (.01, -.015))
    arrow = (*points[2], 0., cos(pi/4), sin(pi/4), 0.)
    visible, tolerance = _radius_arc(arrow, .010, (*arc, tangent))
    assert visible == arc
    assert any(_distance_to_stroke(points[1], edge) <= tolerance for edge in visible)
    assert all(_distance_to_stroke((.01, -.003), edge) > tolerance for edge in visible)


def test_gost_format_reserves_the_upper_left_designation_cell():
    sheet = SimpleNamespace(GetTemplateName="a2 - gost_sh1.slddrt")
    doc = SimpleNamespace(GetCurrentSheet=sheet)
    frame = Frame(.594, .420, True)
    reserved = format_zones(doc, frame)
    assert reserved[0] == pytest.approx((.020, .401, .090, .415))
    box = (.11, .405, .12, .41)
    assert not _routes_fit(box, (((.06, .405), (.11, .405)),), reserved, [])
    assert not fits((.04, .405, .055, .410), frame.drawable, reserved, ())


def test_section_markers_are_reserved_in_sheet_coordinates():
    line = SimpleNamespace(GetLabel="A", GetTextFormat=SimpleNamespace(CharHeight=.007),
                           GetTextInfo=(.050, .228, 0., .169, .228, 0.),
                           GetArrowInfo=(.058, .218, 0., .058, .229, 0., .167, .218, 0., .167, .229, 0.))
    view = SimpleNamespace(GetSectionLines=(line,), GetXform=(.110, .229, 2.))
    zones = tuple(section_line_zones(view))
    assert len(zones) == 4
    assert not fits((.049, .224, .055, .232), (.020, .060, .205, .292), zones, ())
    assert not _routes_fit((.045, .240, .050, .245), (((.05, .224), (.065, .224)),), zones, ())


@pytest.mark.parametrize("width", (.210, .594))
def test_gost_text_cells_require_real_sheet_format_borders(width):
    left, divider, right = width - .125, width - .055, width - .005
    lines = (
        ((left, .005), (left, .060)), ((divider, .045), (divider, .005)),
        ((.020, .005), (right, .005)), ((left, .020), (right, .020)),
        ((left, .045), (right, .045)),
    )
    assert form_cells(width, lines) == {
        "name": (left, .020, divider, .045), "material": (left, .005, divider, .020),
    }
    assert set(form_cells(width, lines[:-1])) == {"material"}
    assert not form_cells(width, lines[2:])


def test_title_block_leaves_linked_notes_in_other_columns_alone():
    inside = SimpleNamespace(GetPosition=(.120, .030, 0.))
    outside = SimpleNamespace(GetPosition=(.190, .030, 0.))
    notes = tuple(SimpleNamespace(PropertyLinkedText='$PRPSHEET:"Наименование"',
                                  GetText="Название", GetAnnotation=annotation)
                  for annotation in (inside, outside))
    drawing = SimpleNamespace(GetFirstView=SimpleNamespace(GetNotes=notes))
    assert _linked_notes(drawing, {"name": (.085, .020, .155, .045)}) == {"name": [inside]}


def test_detail_label_uses_clear_space_inside_a_view_at_its_own_scale():
    polyline = (0, 0, 0, 0, -1, 1, -1, 0, 2, 0, 0, 0, .02, 0, 0)
    strokes = _polyline_strokes(polyline, (.06, .14, 2))
    assert strokes == (((.06, .14), (.1, .14)),)
    box = (.07, .151, .086, .155)
    outline = (.04, .11, .13, .20)
    assert _fits_detail(box, PAPERS[0].drawable, (), (), outline, strokes)
    assert not _fits_detail(box, PAPERS[0].drawable, (), (), outline, ())
    assert not _fits_detail((.07, .138, .086, .142), PAPERS[0].drawable, (), (), outline, strokes)
    assert _polyline_strokes(polyline[:-1], (.06, .14, 2)) == ()
