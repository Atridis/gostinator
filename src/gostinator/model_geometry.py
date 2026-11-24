"""Read analytic solid geometry from the active SolidWorks configuration."""
from collections import defaultdict
from dataclasses import dataclass
from math import pi

from .solidworks import as_dispatch, get_value


# All SolidWorks lengths are in metres. Snap only for grouping coaxial faces.
POSITION_TOLERANCE = 0.000025
NORMAL_TOLERANCE = 0.999
FRONT, BACK, LEFT, RIGHT, TOP, BOTTOM = (
    "*Front", "*Back", "*Left", "*Right", "*Top", "*Bottom"
)
VIEWS = {(0, 1): RIGHT, (0, -1): LEFT, (1, 1): TOP, (1, -1): BOTTOM,
         (2, 1): FRONT, (2, -1): BACK}


def axis_of(direction):
    axis = max(range(3), key=lambda i: abs(direction[i]))
    return axis if abs(direction[axis]) >= NORMAL_TOLERANCE else None


@dataclass(frozen=True)
class Bore:
    axis: int
    center: tuple[float, float]
    radius: float
    low: float
    high: float

    def point(self, depth):
        coords = list(self.center)
        coords.insert(self.axis, depth)
        return tuple(coords)


@dataclass(frozen=True)
class Model:
    bounds: tuple[float, ...]
    bores: tuple[Bore, ...]
    sheet_metal: bool
    outside_loops: tuple[tuple[int, int], ...] = ()

    @property
    def spans(self):
        return tuple(self.bounds[i + 3] - self.bounds[i] for i in range(3))


@dataclass(frozen=True)
class Section:
    along: int
    coordinate: float
    bores: tuple[Bore, ...]


def _read_cylinders(faces):
    segments = defaultdict(list)
    for face in faces:
        surface = get_value(face, "GetSurface")
        if not get_value(surface, "IsCylinder"):
            continue
        params = get_value(surface, "CylinderParams")
        if not params or len(params) < 7:
            continue
        axis = axis_of(params[3:6])
        if axis is None or params[6] <= POSITION_TOLERANCE:
            continue
        # Underlying surface normals can point in either direction. Sample a
        # real point on the cylinder and orient the normal as the solid face.
        uv = get_value(face, "GetUVBounds")
        evaluated = get_value(surface, "Evaluate", (uv[0] + uv[1]) / 2,
                              (uv[2] + uv[3]) / 2, 0, 0)
        if not evaluated or len(evaluated) < 6:
            continue
        others = tuple(i for i in range(3) if i != axis)
        dot = sum((evaluated[i] - params[i]) * evaluated[-3 + i] for i in others)
        if get_value(face, "FaceInSurfaceSense"):
            dot = -dot
        radius = params[6]
        if dot >= -radius * .5:
            continue
        box = get_value(face, "GetBox")
        if not box or box[axis + 3] - box[axis] <= POSITION_TOLERANCE:
            continue
        center = tuple(params[i] for i in others)
        key = (axis, *(round(value / POSITION_TOLERANCE) for value in (
            *center, radius, box[axis], box[axis+3])))
        segments[key].append((Bore(axis, center, radius, box[axis], box[axis + 3]),
                              get_value(face, "GetArea")))
    return tuple(
        rows[0][0] for rows in segments.values()
        if sum(area for _, area in rows) >= .82 * 2 * pi * rows[0][0].radius
        * (rows[0][0].high - rows[0][0].low)
    )


def _outer_loops(faces, bounds):
    counts = [[0, 0] for _ in range(3)]
    for face in faces:
        normal = get_value(face, "Normal")
        if not normal or len(normal) < 3:
            continue
        axis = axis_of(normal)
        if axis is None:
            continue
        side = 1 if normal[axis] > 0 else 0
        box = get_value(face, "GetBox")
        limit = bounds[axis + 3 * side]
        if box and max(abs(box[axis] - limit), abs(box[axis + 3] - limit)) <= POSITION_TOLERANCE:
            counts[axis][side] += len(get_value(face, "GetLoops") or ())
    return tuple(tuple(side) for side in counts)


def read_model(doc):
    bodies = get_value(doc, "GetBodies2", 0, False) or ()
    if not bodies:
        return None
    boxes = [get_value(body, "GetBodyBox") for body in bodies]
    if any(not box or len(box) != 6 for box in boxes):
        return None
    bounds = tuple(min(box[i] for box in boxes) for i in range(3))
    bounds += tuple(max(box[i] for box in boxes) for i in range(3, 6))
    if min(bounds[i+3] - bounds[i] for i in range(3)) <= POSITION_TOLERANCE:
        return None
    # For multi-body models the box is useful for layout, but coincident faces
    # from independent bodies must not be combined into bore evidence.
    if len(bodies) != 1:
        return Model(bounds, (), False)
    body = bodies[0]
    faces = tuple(map(as_dispatch, get_value(body, "GetFaces") or ()))
    return Model(bounds, _read_cylinders(faces), bool(get_value(body, "IsSheetMetal")),
                 _outer_loops(faces, bounds))


def primary_axis(model):
    spans = model.spans
    areas = [spans[(axis + 1) % 3] * spans[(axis + 2) % 3] for axis in range(3)]
    largest = max(areas)
    # The principal face should cover substantial projected area while making
    # the most bore centres measurable. Do not rotate for a single side hole.
    bore_counts = [
        len({bore.center for bore in model.bores if bore.axis == axis}) for axis in range(3)
    ]
    return max((2, 1, 0), key=lambda a: areas[a] / largest + min(bore_counts[a], 8) * .18)


def perpendicular_flange_profile(model):
    if model.sheet_metal:
        axes = {bore.axis for bore in model.bores}
        if len(axes) == 2:
            return next(i for i in range(3) if i not in axes)
    return None


def _coaxial_groups(bores):
    groups = defaultdict(list)
    for bore in bores:
        groups[tuple(round(p / POSITION_TOLERANCE) for p in bore.center)].append(bore)
    return list(groups.values())


def choose_section(model, axis):
    if model.sheet_metal:
        return None
    groups = _coaxial_groups(b for b in model.bores if b.axis == axis)
    if not groups:
        return None
    transverse = tuple(i for i in range(3) if i != axis)
    width = min(model.spans[i] for i in transverse)
    dominant = [
        group for group in groups
        if max(b.radius for b in group) >= .26 * width
        or (len({round(b.radius / POSITION_TOLERANCE) for b in group}) > 1
            and max(b.radius for b in group) >= .15 * width)
    ]
    if not dominant or (len(groups) > 8 and len(dominant) < len(groups) / 3):
        return None
    candidates = []
    for cross in (0, 1):
        along = transverse[1 - cross]
        for exemplar in dominant:
            coordinate = exemplar[0].center[cross]
            intersected = [group for group in groups
                           if abs(group[0].center[cross] - coordinate) < POSITION_TOLERANCE]
            score = (sum(g in dominant for g in intersected), len(intersected), model.spans[along])
            candidates.append((score, Section(along, coordinate, tuple(b for group in intersected for b in group))))
    return max(candidates, key=lambda row: row[0])[1]


def needs_reverse(model, axis):
    # Through and blind bore entry points alone miss pockets and recesses.
    # Compare the number of actual inner/outer loops on the two broad faces.
    if model.outside_loops and model.spans[axis] <= .16 * min(
            model.spans[i] for i in range(3) if i != axis):
        low, high = model.outside_loops[axis]
        if min(low, high) >= 4 and abs(low - high) >= max(4, min(low, high) * .6):
            return True
    groups = _coaxial_groups(b for b in model.bores if b.axis == axis)
    if len(groups) < 4:
        return False
    clearance = max(POSITION_TOLERANCE * 2, .07 * model.spans[axis])
    low, high = model.bounds[axis], model.bounds[axis + 3]
    distinct = 0
    for group in groups:
        front = {round(b.radius / POSITION_TOLERANCE) for b in group if high - b.high < clearance}
        reverse = {round(b.radius / POSITION_TOLERANCE) for b in group if b.low - low < clearance}
        distinct += bool(reverse - front)
    return distinct >= 2
