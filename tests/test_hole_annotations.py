from types import SimpleNamespace

from gostinator.annotation_layout import _clear_leader, _leader_positions
from gostinator.drawing_dimensions import Anchor, Circle
from gostinator.hole_annotations import _groups, _linear_diameter_bore, _pitch_circle, axis_connections, projected_holes
from gostinator.model_geometry import Bore, Model


MATRIX = (1., 0., 0., 0., 1., 0., 0., 0., 1., .10, .10, 0., 1., 0., 0., 0.)


def _circle(x, y, z, radius):
    return Circle(Anchor(object(), (x, y, z), (x+.10, y+.10, z)), 2, radius, 1.)


def test_same_diameter_with_a_different_depth_or_thread_is_a_different_group():
    circles = [_circle(x, 0., .01, .002) for x in (0., .02, .04, .06, .08)]
    bores = [Bore(2, (x, 0.), .002, .004 if x == .04 else 0., .01)
             for x in (0., .02, .04, .06, .08)]
    threads = (((.16, .10), ('M5', 200, 400, '6H')),)
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


def test_leader_route_allows_its_rim_but_rejects_line_crossings_and_overlap():
    route = (((0., 0.), (.04, 0.)),)
    rim = (((.003, -.001), (.003, .001)),)
    cross = (((.025, -.006), (.025, .006)),)
    overlap = (((.02, 0.), (.05, 0.)),)
    assert _clear_leader(route, rim, (0., 0.), .0032)
    assert not _clear_leader(route, rim + cross, (0., 0.), .0032)
    assert not _clear_leader(route, overlap, (0., 0.), .0032)
