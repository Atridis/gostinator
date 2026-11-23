from types import SimpleNamespace

import pytest

from gostinator.drawing_dimensions import (
    Anchor, Line, _corner_cut_tasks, _overall_tasks, _section_frame, _select_section_edge,
)
from gostinator.drawing_views import project
from gostinator.model_geometry import Bore, Model, _read_cylinders, choose_section, needs_reverse, primary_axis


def _bore(center, radius, low, high, axis=2):
    return Bore(axis, center, radius, low, high)


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
