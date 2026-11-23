from math import pi
from pathlib import Path
from types import SimpleNamespace

import pytest

from gostinator.annotation_layout import _routes_fit, candidates, crosses, fits, format_zones, text_extent
from gostinator.drawing_views import Frame
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


def test_gost_format_reserves_the_upper_left_designation_cell():
    sheet = SimpleNamespace(GetTemplateName="a2 - gost_sh1.slddrt")
    doc = SimpleNamespace(GetCurrentSheet=sheet)
    frame = Frame(.594, .420, True)
    reserved = format_zones(doc, frame)
    assert reserved[0] == pytest.approx((.020, .401, .090, .415))
    box = (.11, .405, .12, .41)
    assert not _routes_fit(box, (((.06, .405), (.11, .405)),), reserved, [])
    assert not fits((.04, .405, .055, .410), frame.drawable, reserved, ())
