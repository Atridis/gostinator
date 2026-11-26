from math import cos, sin, tau
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gostinator import hole_annotations
from gostinator.annotation_layout import DimensionTextLayout, _clear_leader, _leader_positions
from gostinator.drawing_dimensions import Anchor, Circle
from gostinator.hole_annotations import ThreadSpec, _groups, _linear_diameter_bore, _pitch_circle, axis_connections
from gostinator.hole_annotations import projected_holes
from gostinator.model_geometry import Bore, Model


MATRIX = (1., 0., 0., 0., 1., 0., 0., 0., 1., .10, .10, 0., 1., 0., 0., 0.)


def _circle(x, y, z, radius):
    return Circle(Anchor(object(), (x, y, z), (x+.10, y+.10, z)), 2, radius, 1.)


def test_same_diameter_with_a_different_depth_or_thread_is_a_different_group():
    circles = [_circle(x, 0., .01, .002) for x in (0., .02, .04, .06, .08)]
    bores = [Bore(2, (x, 0.), .002, .004 if x == .04 else 0., .01)
             for x in (0., .02, .04, .06, .08)]
    threads = (((.16, .10), ThreadSpec('M5', 200, 400, '6H')),)
    model = Model((0., -.02, 0., .10, .02, .01), tuple(bores), False)
    holes = projected_holes(circles, model, MATRIX, threads)

    group, = _groups(holes)
    assert {hole.bores[0].center[0] for hole in group} == {0., .02, .08}
    assert len(tuple(axis_connections(holes, 1.))) == 1  # no line through the other two holes


def test_equal_stepped_holes_have_one_center_each_and_share_axis_only_when_aligned():
    positions = ((0., 0.), (.020, 0.), (.034, .012))
    bores = tuple(Bore(2, pos, r, low, high) for pos in positions
                  for r, low, high in ((.002, 0., .005), (.004, .005, .010)))
    circles = [_circle(x, y, b.high, b.radius) for b in bores for x, y in (b.center,)]
    circles.append(_circle(.034, .026, .010, .003))  # no physical bore

    holes = projected_holes(circles, Model((0., 0., 0., .05, .03, .01), bores, False), MATRIX)

    group, = _groups(holes)
    assert len(group) == 3 and all(hole.rim.radius == .004 for hole in group)
    assert len(tuple(axis_connections(holes, 1.))) == 1


def test_a_bolt_circle_needs_regular_angles_and_a_common_center_bore():
    points = ((-.020, 0.), (0., -.020), (.020, 0.), (0., .020), (0., 0.))
    bores = tuple(Bore(2, point, .002 if i < 4 else .008, 0., .01)
                  for i, point in enumerate(points))
    circles = [_circle(*b.center, b.high, b.radius) for b in bores]
    model = Model((-.03, -.03, 0., .03, .03, .01), bores, False)
    holes = projected_holes(circles, model, MATRIX)
    group, = _groups(holes)
    center, radius = _pitch_circle(group, holes, 1.)
    assert abs(center[0]-.10) < 1e-9 and abs(center[1]-.10) < 1e-9
    assert abs(radius-.020) < 1e-9
    assert _pitch_circle(group, group, 1.) is None
    assert _pitch_circle(group[:3], holes, 1.) is None


def test_offset_diameter_is_matched_by_both_extension_lines():
    bore = Bore(2, (.02, .02), .004, 0., .01)
    other = Bore(2, (.02, .032), .004, 0., .01)
    tips = ((.116, .16), (.124, .16))
    strokes = [(0, x, .12, 0., x, .161, 0.) for x, _ in tips]
    data = SimpleNamespace(GetArrowHeadCount=2, GetArrowHeadAtIndex2=tips.__getitem__,
                           GetLineCount=2, GetLineAtIndex=strokes.__getitem__)
    annotation = SimpleNamespace(GetDisplayData=data)

    assert _linear_diameter_bore(annotation, .008, data, MATRIX, (bore, other), 2) == bore
    assert _linear_diameter_bore(annotation, .008, data, MATRIX, (other,), 2) is None


def test_free_route_can_select_another_hole_of_the_same_family():
    anchors = (("crowded", (.04, .04), .003), ("clear", (.12, .12), .003))
    wall = (((.025, .025), (.055, .025)), ((.055, .025), (.055, .055)),
            ((.055, .055), (.025, .055)), ((.025, .055), (.025, .025)))
    occupied = ((.023, .023, .057, .057),)
    sheet = (0., 0., .2, .2)
    candidates = _leader_positions(anchors, .024, .006, sheet, sheet, (), occupied, wall)
    assert candidates and {row[2] for row in candidates} == {"clear"}


def test_clear_route_prefers_a_center_without_a_callout_and_can_fall_back():
    view = SimpleNamespace(GetName2="End")
    occupied, free = (.04, .04), (.12, .12)
    anchors = (("occupied", occupied, .003), ("free", free, .003))
    layout = object.__new__(DimensionTextLayout)
    layout.frame = (0., 0., .2, .2)
    layout.boxes = {"End": (.02, .02, .16, .16)}
    layout.edges = {"End": (((.02, .02), (.16, .02)),)}
    layout.reserved = []
    layout.occupied = {}

    routes = list(layout.free_diameter_routes(view, None, anchors, (.024, .006), (), used_centers={occupied}))

    assert routes[0][2] == free
    assert routes[-1][2] == occupied


def test_layout_keeps_same_short_dimension_names_from_different_features_separate():
    view = SimpleNamespace(GetName2="End")
    annotations = [SimpleNamespace(GetName="D2", GetSpecificAnnotation=SimpleNamespace(
        GetDimension2=lambda _, feature=feature: SimpleNamespace(FullName=f"D2@{feature}@Part.Part")))
        for feature in ("Spacing", "Counterbore")]

    assert len({DimensionTextLayout._key(view, annotation) for annotation in annotations}) == 2


@pytest.mark.parametrize("other_center_fits", (True, False))
def test_existing_diameter_can_move_to_an_unused_center_or_keep_its_place(monkeypatch, other_center_fits):
    bores = tuple(Bore(2, (x, 0.), .002, 0., .01) for x in (0., .02))
    model = Model((0., -.02, 0., .03, .02, .01), bores, False)
    holes = projected_holes([_circle(*b.center, b.high, b.radius) for b in bores], model, MATRIX)
    view = SimpleNamespace(GetName2="End")
    original, replacement = object(), object()
    item = SimpleNamespace(view=view, annotation=original, dimension=SimpleNamespace(FullName="RD1@End.Drawing"))
    route = Mock(return_value=replacement if other_center_fits else None)
    reuse = Mock(return_value=True)
    monkeypatch.setattr(hole_annotations, "_route_group_diameter", route)
    monkeypatch.setattr(hole_annotations, "_place_quantity", reuse)
    doc, layout = object(), object()
    used = {holes[0].center}

    result = hole_annotations._place_view_diameter(doc, layout, item, holes, bores[0], MATRIX, holes, used)

    route.assert_called_once_with(doc, layout, view, item, holes, bores[0], MATRIX, holes, used_centers=used)
    assert result is (replacement if other_center_fits else original)
    assert reuse.call_count == int(not other_center_fits)


def test_leader_route_allows_its_rim_but_rejects_line_crossings_and_overlap():
    route = (((0., 0.), (.04, 0.)),)
    rim = (((.003, -.001), (.003, .001)),)
    cross = (((.025, -.006), (.025, .006)),)
    overlap = (((.02, 0.), (.05, 0.)),)
    assert _clear_leader(route, rim, (0., 0.), .0032)
    assert not _clear_leader(route, rim + cross, (0., 0.), .0032)
    assert not _clear_leader(route, overlap, (0., 0.), .0032)


def test_external_leader_can_cross_one_proven_outline_edge_only():
    route = (((0., 0.), (.04, 0.)),)
    outside = ((.025, -.006), (.025, .006))
    inside = ((.016, -.006), (.016, .006))
    assert _clear_leader(route, (outside,), (0., 0.), .003, border=(outside,))
    assert not _clear_leader(route, (inside, outside), (0., 0.), .003, border=(outside,))
    assert not _clear_leader(route, (inside, outside), (0., 0.), .003, border=(inside, outside))


def test_thread_diameter_arrowheads_must_bracket_the_collected_hole_center():
    center = (.10, .08)
    tips = ((.097, .08), (.103, .08))
    data = SimpleNamespace(GetArrowHeadCount=2, GetArrowHeadAtIndex2=tips.__getitem__)
    annotation = SimpleNamespace(GetDisplayData=data)

    assert hole_annotations._arrows_span_hole(annotation, center, .003)
    assert not hole_annotations._arrows_span_hole(annotation, (.10, .081), .003)
    assert not hole_annotations._arrows_span_hole(annotation, center, .005)
    data.GetArrowHeadCount = 1
    assert not hole_annotations._arrows_span_hole(annotation, center, .003)


def test_finer_leader_directions_can_pass_between_close_obstacles():
    center = (.10, .10)
    anchors = (("hole", center, .003),)
    directions = [(cos(i*tau/8), sin(i*tau/8)) for i in range(8)]
    obstacles = tuple(tuple((center[0]+.004*dx-sign*.00025*dy,
                             center[1]+.004*dy+sign*.00025*dx) for sign in (-1, 1))
                      for dx, dy in directions)
    sheet = (0., 0., .2, .2)
    args = (anchors, .020, .004, sheet, sheet, (), (), obstacles)

    assert not _leader_positions(*args)
    assert _leader_positions(*args, refine=True)


@pytest.mark.parametrize("end_result", ("placed", "crowded", "not_visible"))
def test_section_diameter_only_supplies_a_quantity_for_an_end_view(monkeypatch, end_result):
    cut = SimpleNamespace(GetName2="A-A", Type=2)
    end = SimpleNamespace(GetName2="End", Type=7)
    data = SimpleNamespace(GetTextCount=1, GetTextAtIndex=lambda _: "<MOD-DIAM>4")
    annotation = SimpleNamespace(GetName="D1", GetDisplayData=data)
    item = SimpleNamespace(view=cut, annotation=annotation, dimension=object())
    bores = tuple(Bore(2, (x, 0.), .002, 0., .01) for x in (0., .02))
    model = Model((0., -.02, 0., .03, .02, .01), bores, False)
    circles = [_circle(*b.center, b.high, b.radius) for b in bores]
    holes = projected_holes(circles, model, MATRIX)
    layout = SimpleNamespace(items=(item,))
    monkeypatch.setattr(hole_annotations, "DimensionTextLayout", Mock(return_value=layout))
    route = Mock(return_value=object() if end_result == "placed" else None)
    monkeypatch.setattr(hole_annotations, "_route_group_diameter", route)
    prefix = Mock()
    monkeypatch.setattr(hole_annotations, "_place_quantity", prefix)
    doc, frame = object(), object()
    views = [(cut, holes, MATRIX, False), (end, holes, MATRIX, end_result == "not_visible")]

    result = hole_annotations.add_hole_quantities(doc, frame, views, [(bores[0], annotation)])

    assert result == (int(end_result == "placed"), 0)
    prefix.assert_not_called()
    if end_result == "not_visible":
        route.assert_not_called()
    else:
        route.assert_called_once_with(doc, layout, end, item, holes, bores[0], MATRIX, holes, used_centers=set())


@pytest.mark.parametrize("note_placed", (True, False))
def test_grouped_thread_hides_only_its_generated_drill_diameter(monkeypatch, note_placed):
    view = SimpleNamespace(GetName2="End", Type=7)
    drills = [Bore(2, (x, 0.), .0021, 0., .008) for x in (0., .02, .04)]
    counterbores = [Bore(2, (x, 0.), .003, .008, .01) for x in (0., .02)]
    bores = (*drills, *counterbores)
    circles = [_circle(*b.center, b.high, b.radius) for b in bores]
    spec = ThreadSpec("M5", 200, 400, "6H", tap_drill=168)
    threads = (((.10, .10), spec), ((.12, .10), spec))
    model = Model((0., -.02, 0., .05, .02, .01), bores, False)
    holes = projected_holes(circles, model, MATRIX, threads)
    assert [hole.tap_drill for hole in holes] == [drills[0], drills[1], None]

    labels = {"thread": (drills[0], "M5"), "drill": (drills[0], ""), "plain": (drills[2], ""),
              "counterbore": (counterbores[0], ""), "explicit": (drills[1], "")}
    items = {
        name: SimpleNamespace(view=view, kind=6, display=object(),
                              annotation=SimpleNamespace(GetName=name, Visible=1),
                              dimension=SimpleNamespace(FullName=("D1@Part.Part" if name in ("thread", "explicit")
                                                                  else "RD1@Sheet.Drawing")))
        for name in labels
    }
    layout = SimpleNamespace(items=tuple(items.values()), occupied={},
                             _key=lambda view, annotation: (view.GetName2, annotation.GetName))
    monkeypatch.setattr(hole_annotations, "DimensionTextLayout", Mock(return_value=layout))
    monkeypatch.setattr(hole_annotations, "_diameter_label",
                        lambda display, annotation, dimension, matrix, bores: labels[annotation.GetName])
    monkeypatch.setattr(hole_annotations, "_route_group_diameter", Mock(return_value=None))
    monkeypatch.setattr(hole_annotations, "_place_thread_group",
                        Mock(return_value=object() if note_placed else None))

    result = hole_annotations.add_hole_quantities(object(), object(), [(view, holes, MATRIX, False)], ())

    assert result == (int(note_placed), int(note_placed))
    visible_changes = {name for name, item in items.items() if item.annotation.Visible == 3}
    assert visible_changes == ({"drill"} if note_placed else set())
