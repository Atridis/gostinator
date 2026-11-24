from math import pi
from types import SimpleNamespace

import pytest

from gostinator.drawing_dimensions import (
    Anchor, Circle, Line, Task, _arc_vertex_shift, _blank_extent_tasks, _closed_inner_contours, _corner_cut_tasks,
    _cut_depth, _depth_extends_past_material, _folded_profile_shift, _geometry_key,
    _overall_tasks, _physical_profile_key, _recess_tasks, _section_axis_lines,
    _section_datums, _section_depth_task, _section_frame, _select_section_edge, _stock_span_tasks,
    prune_duplicate_bore_diameters,
)
from gostinator.drawing_views import _flat_plane, project
from gostinator.model_geometry import Bore, Model, Section, _read_cylinders, choose_section, needs_reverse, primary_axis


def _bore(center, radius, low, high, axis=2):
    return Bore(axis, center, radius, low, high)


def test_cut_depth_requires_a_subtractive_feature_and_matching_depth_parameter():
    definition = SimpleNamespace(BothDirections=True, GetDepth=lambda forward: .02 if forward else .006,
                                 GetEndCondition=lambda forward: 0)
    cut = SimpleNamespace(GetTypeName2="Cut", GetDefinition=definition)
    boss = SimpleNamespace(GetTypeName2="Boss", GetDefinition=definition)
    imported = SimpleNamespace(GetFeatureOwner=cut, GetSystemValue2=lambda *_: .006)
    stock = SimpleNamespace(GetFeatureOwner=boss, GetSystemValue2=lambda *_: .02)
    other_cut_parameter = SimpleNamespace(GetFeatureOwner=cut, GetSystemValue2=lambda *_: .004)
    assert _cut_depth(None, imported, "D2@Cut@Piece.Part") == .006
    assert _cut_depth(None, stock, "D1@Boss@Piece.Part") is None
    assert _cut_depth(None, other_cut_parameter, "D3@Cut@Piece.Part") is None
    assert _cut_depth(None, imported, "RD1@Section@Sheet.Drawing") is None


def test_cut_length_checks_the_real_material_end_and_keeps_depths_inside_stock():
    contour = (((.05, .10), (.15, .10)), ((.15, .10), (.15, .115)),
               ((.15, .115), (.05, .115)), ((.05, .115), (.05, .10)))
    assert _depth_extends_past_material(((.18, .10), (.18, .12)), contour, .02, 1.)
    assert _depth_extends_past_material(((.18, .115), (.18, .095)), contour, .02, 1.)
    assert not _depth_extends_past_material(((.18, .10), (.18, .115)), contour, .015, 1.)
    assert not _depth_extends_past_material(((.18, .115), (.18, .109)), contour, .006, 1.)
    assert not _depth_extends_past_material(((.18, .10), (.18, .12)), contour, .02, 2.)
    assert not _depth_extends_past_material(((.18, .10), (.18, .12)), (), .02, 1.)
    assert not _depth_extends_past_material(((.18, .130), (.18, .150)), contour, .02, 1.)


def test_sheet_blank_plane_is_validated_against_its_actual_thickness():
    face = SimpleNamespace(Normal=(0., 1., 0.), GetArea=.002)

    def shape(height):
        vertices = (SimpleNamespace(GetPoint=(0., 0., 0.)),
                    SimpleNamespace(GetPoint=(.101, height, .024)))
        body = SimpleNamespace(GetFaces=(face,), GetVertices=vertices)
        return SimpleNamespace(GetBodies2=lambda *_: (body,))

    assert _flat_plane(shape(.008), .008) == (0., 1., 0.)
    assert _flat_plane(shape(.052), .008) is None


def test_view_choice_prioritizes_a_row_of_real_bores():
    model = Model((-.040, -.010, 0, .040, .010, .022), (
        _bore((-.021, 0), .003, 0, .013), _bore((-.021, 0), .007, .013, .022),
        _bore((.021, 0), .003, 0, .013), _bore((.021, 0), .007, .013, .022),
    ), False)
    assert primary_axis(model) == 2
    cut = choose_section(model, 2)
    assert cut is not None
    assert cut.along == 0 and cut.coordinate == 0 and len(cut.bores) == 4

    rotated = Model((-.010, -.040, 0, .010, .040, .022), tuple(
        _bore(bore.center[::-1], bore.radius, bore.low, bore.high) for bore in model.bores), False)
    assert primary_axis(rotated) == 2
    assert choose_section(rotated, 2).along == 1
    assert choose_section(Model(model.bounds, model.bores, True), 2) is None


def test_section_axes_cover_coaxial_steps_and_closed_slots_only():
    cut = Section(0, 0., (
        _bore((.010, 0.), .002, .0005, .006),
        _bore((.010, 0.), .003, .006, .0115),
    ))
    matrix = (1, 0, 0, 0, 0, 1, 0, 1, 0, .10, .10, 0, 1, 0, 0, 0)
    base_matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, .10, .10, 0, 1, 0, 0, 0)
    hatches = tuple((lo, 0., 0., hi, 0., .012)
                    for lo, hi in ((0., .008), (.012, .018), (.022, .028), (.032, .040)))
    closed = ((.118, .098), (.122, .098), (.122, .102), (.118, .102), (.118, .098))
    opened = ((.128, .098), (.132, .098), (.132, .102), (.128, .102))
    strokes = [*zip(closed, closed[1:]), *zip(opened, opened[1:])]
    contours = tuple(_closed_inner_contours(strokes, (.10, .09, .14, .11)))

    lines = _section_axis_lines(cut, matrix, (0., 0., 0.), hatches, base_matrix, contours)

    assert len(lines) == 2
    for axis_line, x in zip(lines, (.110, .120)):
        a, b = axis_line
        assert a == pytest.approx((x, .098, 0.))
        assert b == pytest.approx((x, .114, 0.))


@pytest.mark.parametrize(("same_profile", "cut_visible", "extra_text", "removed"), [
    (True, True, "", 1),
    (False, True, "", 0),
    (True, False, "", 0),
    (True, True, " H7", 0),
])
def test_section_diameter_only_replaces_a_plain_projection_of_the_same_step(
        same_profile, cut_visible, extra_text, removed):
    left = _bore((-.021, 0), .003, 0, .013)
    right = _bore((.021, 0), .003, 0, .013)
    model = Model((-.040, -.010, 0, .040, .010, .022), (
        left, right, _bore(left.center, .007, .013, .022),
        _bore(right.center, .007 if same_profile else .008, .013, .022),
    ), False)
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, .10, .12, 0, 2., 0, 0, 0)
    center = project(right.point(right.low), matrix)
    text = ("<MOD-DIAM>", "6", extra_text)
    data = SimpleNamespace(
        GetTextCount=len(text), GetTextAtIndex=text.__getitem__,
        GetArrowHeadCount=2,
        GetArrowHeadAtIndex2=lambda i: (center[0] + (2*i - 1)*.006, *center[1:]),
    )
    front = SimpleNamespace(Visible=1, IsDangling=False, GetDisplayData=data)
    cut = SimpleNamespace(Visible=1 if cut_visible else 3, IsDangling=False, GetDisplayData=data)
    dimension = SimpleNamespace(GetSystemValue2=lambda _: .006)
    display = SimpleNamespace(Type2=6, GetAnnotation=front, GetDimension2=lambda _: dimension)
    view = SimpleNamespace(Type=7, GetDisplayDimensions=(display,),
                           ModelToViewTransform=SimpleNamespace(ArrayData=matrix))
    doc = SimpleNamespace(GetViews=((None, view),))

    assert prune_duplicate_bore_diameters(doc, model, [(left, cut)]) == removed
    assert front.Visible == (3 if removed else 1)
    assert cut.Visible == (1 if cut_visible else 3)


def test_opposite_face_only_for_distinct_backside_openings():
    bounds = (-.03, -.02, 0, .03, .02, .012)
    through = tuple(_bore((i * .008, .01), .002, 0, .012) for i in (0, 1, 2, 3))
    assert not needs_reverse(Model(bounds, through, False), 2)
    backside = (_bore((0, -.01), .003, 0, .005), _bore((.008, -.01), .003, 0, .005))
    assert needs_reverse(Model(bounds, through + backside, False), 2)


def test_recessed_outer_faces_can_justify_an_opposite_view():
    bounds = (-.08, -.06, 0, .08, .06, .008)
    asymmetric = ((1, 1), (1, 1), (7, 14))
    assert needs_reverse(Model(bounds, (), False, asymmetric), 2)
    assert not needs_reverse(Model(bounds, (), False, ((1, 1), (1, 1), (9, 10))), 2)


def test_analytic_outer_radius_and_incomplete_slot_wall_are_not_bores():
    def cylinder(sense, fraction):
        surface = SimpleNamespace(
            IsCylinder=lambda: True,
            CylinderParams=(0., 0., 0., 0., 0., 1., .003),
            Evaluate=lambda *_: (.003, 0., .005, 1., 0., 0.),
        )
        return SimpleNamespace(
            GetSurface=lambda: surface,
            GetUVBounds=lambda: (0., 1., 0., 1.),
            FaceInSurfaceSense=lambda: sense,
            GetBox=lambda: (-.003, -.003, 0., .003, .003, .01),
            GetArea=lambda: fraction * 2 * 3.141592653589793 * .003 * .01,
        )

    assert not _read_cylinders((cylinder(False, 1),))
    assert not _read_cylinders((cylinder(True, .50),))
    assert len(_read_cylinders((cylinder(True, .50), cylinder(True, .50)))) == 1


def test_overall_dimension_uses_axis_projection_for_staggered_vertices():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, .5, 0, 0, 0)
    a = Anchor(object(), (0, 0, 0), (0, 0, 0))
    b = Anchor(object(), (.05, 0, .01), (.025, 0, .005))
    c = Anchor(object(), (.05, 0, 0), (.025, 0, 0))
    projected = list(_overall_tasks([a, b], matrix))[0]
    assert projected.nominal == pytest.approx(.05)
    assert projected.projected
    assert list(_overall_tasks([a, b, c], matrix))[0].nominal == pytest.approx(.05)


def test_blank_dimensions_only_cover_the_confirmed_contour_extents():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, .5, 0, 0, 0)
    points = ((0., 0., 0.), (.080, 0., 0.), (.080, .040, 0.), (0., .040, 0.))
    anchors = [Anchor(object(), point, project(point, matrix)) for point in points]
    assert not list(_blank_extent_tasks(anchors, matrix, None))
    tasks = list(_blank_extent_tasks(anchors, matrix, (0., 0., .040, .020)))
    assert [task.nominal for task in tasks] == pytest.approx([.080, .040])
    assert all(task.overall and task.projected for task in tasks)
    task, = _blank_extent_tasks(anchors, matrix, (0., 0., .042, .020))
    assert task.direction == 1  # the vertices don't reach the full width


def test_rounded_blank_uses_native_arc_attachments_for_its_overall_length():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    anchors = [Anchor(object(), p, p) for p in (
        (.012, 0., 0.), (.088, 0., 0.), (.012, .024, 0.), (.088, .024, 0.))]
    ends = [Circle(Anchor(object(), p, p), 2, .012, .50)
            for p in ((.012, .012, 0.), (.088, .012, 0.))]
    tasks = list(_blank_extent_tasks(anchors, matrix, (0., 0., .100, .024), ends))
    assert {(task.direction, round(task.nominal, 6), task.arc_ends) for task in tasks} == {
        (0, .100, 2), (1, .024, 0)}


def test_opposite_views_deduplicate_the_same_span_but_keep_an_independent_span():
    front = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    back = (-1, 0, 0, 0, 1, 0, 0, 0, -1, 0, 0, 0, 1, 0, 0, 0)

    def measure(y, z):
        a = Anchor(object(), (.01, y, z), (.01, y, z))
        b = Anchor(object(), (.03, y, z), (.03, y, z))
        return Task((a, b), 2, .02, 0)

    original = _geometry_key(measure(.02, .01), front, 2)
    assert original == _geometry_key(measure(.02, 0), back, 2)
    assert original != _geometry_key(measure(.03, 0), back, 2)
    assert original != _geometry_key(measure(.02, 0), back, 2, "Flat pattern")


def test_perpendicular_views_recognize_the_same_physical_profile_edge_after_registration():
    front = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    top = (1, 0, 0, 0, 0, 1, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0)

    def profile(matrix, depth, offset):
        points = ((.002, 0., depth), (.021, 0., depth))
        points = [tuple(a + b for a, b in zip(point, offset)) for point in points]
        anchors = tuple(Anchor(object(), point, project(point, matrix)) for point in points)
        return Task(anchors, 2, .019, 0, profile=True)

    side_offset = (.10, .22, 0.)
    top_offset = (.10, .14, 0.)
    key = _physical_profile_key(profile(front, .05, side_offset), front, side_offset)
    assert key == _physical_profile_key(profile(top, .05, top_offset), top, top_offset)
    assert key != _physical_profile_key(profile(top, .04, top_offset), top, top_offset)


def test_folded_view_registration_checks_the_available_stock_extents():
    model = Model((0., 0., 0., .110, .040, .050), (), True)
    shifted = (Anchor(None, (.112, .170, .02), ()),
               Anchor(None, (.222, .210, .05), ()))
    assert _folded_profile_shift(shifted, model, 2, (.112, .170)) == (.112, .170, 0.)
    assert _folded_profile_shift(shifted, model, 2, (.112, .182)) is None


def test_sheet_outline_registration_matches_raw_arc_ends_to_shifted_vertices():
    points = ((.015, -.012, 0.), (.015, .012, 0.))
    circle = Circle(Anchor(None, (.015, 0., 0.), ()), 2, .012, .5, points)
    shifted = [Anchor(None, (x+.11, y+.17, z), ()) for x, y, z in points]
    unshifted = [Anchor(None, point, ()) for point in points]

    assert _arc_vertex_shift(shifted, [circle], (.11, .17)) == (.11, .17, 0.)
    assert _arc_vertex_shift(unshifted, [circle], (.11, .17)) == (0., 0., 0.)
    assert _arc_vertex_shift(shifted, [circle], (.11, .15)) is None


def test_flat_width_prefers_opposite_parallel_edges_over_tangency_vertices():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    left, right = [Line(object(), tuple(Anchor(object(), p, p) for p in ((x, 0., 0.), (x, .06, 0.))))
                   for x in (0., .024)]
    width, = _blank_extent_tasks(left.ends+right.ends, matrix, (0., -.012, .024, .072),
                                 lines=(left, right))
    assert width.nominal == pytest.approx(.024)
    assert tuple(a.entity for a in width.targets) == (left.entity, right.entity)


def test_corner_cutback_needs_two_perpendicular_incident_rails():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    a = Anchor(object(), (0., .01, 0.), (0., .01, 0.))
    b = Anchor(object(), (.012, 0., 0.), (.012, 0., 0.))
    y = Anchor(object(), (0., .03, 0.), (0., .03, 0.))
    x = Anchor(object(), (.042, 0., 0.), (.042, 0., 0.))
    diagonal = Line(object(), (a, b))
    vertical = Line(object(), (a, y))
    horizontal = Line(object(), (b, x))
    assert not list(_corner_cut_tasks((diagonal, vertical), matrix))
    tasks = list(_corner_cut_tasks((diagonal, vertical, horizontal), matrix))
    assert [task.nominal for task in tasks] == pytest.approx([.012, .01])
    assert all(task.projected for task in tasks)


def test_symmetric_recess_uses_floor_wall_and_rim_connections():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    coordinates = [(0., .01, 0.), (.01, .01, 0.), (.02, 0., 0.),
                   (.04, 0., 0.), (.05, .01, 0.), (.06, .01, 0.)]
    anchors = [Anchor(object(), point, point) for point in coordinates]
    lines = [Line(object(), (a, b)) for a, b in zip(anchors, anchors[1:])]
    assert not list(_recess_tasks(lines[:-1], matrix))
    tasks = list(_recess_tasks(lines, matrix))
    assert [task.kind for task in tasks] == [2, 2, 3]
    assert [task.nominal for task in tasks] == pytest.approx([.01, .02, pi/2])
    assert tasks[0].targets == (anchors[2], anchors[1])
    assert [a.entity for a in tasks[2].targets] == [lines[1].entity, lines[3].entity]


def test_stock_span_requires_an_inner_coplanar_boundary_at_the_same_end():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)

    def rail(y, start, end, z=0.):
        points = ((start, y, z), (end, y, z))
        return Line(object(), tuple(Anchor(object(), p, p) for p in points))

    bottom = rail(0., 0., .06)
    datum = rail(.02, 0., .06)
    partial = rail(.04, 0., .035)
    unrelated = rail(.07, 0., .04, z=.01)
    top = rail(.08, 0., .06)
    task, = _stock_span_tasks((bottom, datum, partial, unrelated, top), matrix)
    assert task.kind == 2 and task.direction == 1 and task.projected
    assert task.nominal == pytest.approx(.02)
    assert [a.entity for a in task.targets] == [datum.entity, partial.entity]


def test_section_hit_must_match_the_selected_curve_and_view():
    matrix = (1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    view = SimpleNamespace(GetName2="Section", GetOutline=(0., 0., .1, .1),
                           ModelToViewTransform=SimpleNamespace(ArrayData=matrix))
    edge = SimpleNamespace(
        GetCurve=SimpleNamespace(IsLine=True),
        GetStartVertex=SimpleNamespace(GetPoint=(.025, .010, 0.)),
        GetEndVertex=SimpleNamespace(GetPoint=(.025, .080, 0.)),
    )
    manager = SimpleNamespace(GetSelectedObject6=lambda *_: edge,
                              GetSelectedObjectsDrawingView2=lambda *_: view)
    doc = SimpleNamespace(Extension=SimpleNamespace(SelectByID2=lambda *_: True))
    assert _select_section_edge(doc, view, manager, Anchor(None, (0, 0, 0), (.025, .040, 0.)), 0)
    assert not _select_section_edge(doc, view, manager, Anchor(None, (0, 0, 0), (.027, .040, 0.)), 0)
    manager.GetSelectedObjectsDrawingView2 = lambda *_: SimpleNamespace(GetName2="Other view")
    assert not _select_section_edge(doc, view, manager, Anchor(None, (0, 0, 0), (.025, .040, 0.)), 0)


def test_section_body_translation_is_corrected_using_its_native_contour():
    model = Model((-.030, -.010, 0., .030, .010, .016),
                  (_bore((0., 0.), .012, 0., .016),), False)
    matrix = (1, 0, 0, 0, 0, -1, 0, 1, 0, .07, .1, 0, 1, 0, 0, 0)
    vertices = [Anchor(object(), (x + .14, .085, z), (0, 0, 0))
                for x in (-.030, .030) for z in (0., .016)]
    effective, shift = _section_frame(vertices, model, 1, matrix, (.04, .1, .10, .116))
    assert shift == pytest.approx((.14, .085, 0.))
    assert project(vertices[-1].model, effective)[:2] == pytest.approx((.1, .116))


@pytest.mark.parametrize("axis,sign", ((0, 1), (1, -1)))
def test_counterbore_depth_measures_the_face_and_shoulder_across_chamfers(axis, sign):
    coords = [(.000, .008), (.000, .004), (.0002, .0038), (.0075, .0038),
              (.008, .0043), (.008, .0055), (.0177, .0055), (.018, .0058), (.018, .009)]
    anchors = []
    for axial, radial in coords:
        point = (sign*axial, radial, 0.) if axis == 0 else (radial, sign*axial, 0.)
        anchors.append(Anchor(object(), point, point))
    lines = tuple(Line(object(), (a, b)) for a, b in zip(anchors, anchors[1:]))
    datums = _section_datums(lines, axis, 1.)
    task = _section_depth_task(anchors[5:7], 0., 1-axis, 1., datums)
    assert task.projected and task.direction == axis
    assert task.nominal == pytest.approx(.010)
    assert task.targets == (anchors[5], anchors[7])
    assert _section_depth_task(anchors[2:4], 0., 1-axis, 1., datums) is None

    missing_face = _section_datums(lines[:-1], axis, 1.)
    assert _section_depth_task(anchors[5:7], 0., 1-axis, 1., missing_face) is None
