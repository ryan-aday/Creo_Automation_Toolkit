from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Iterable

import numpy as np
import trimesh

from .models import BodyRecord, GapRule


@dataclass(slots=True)
class PairDistance:
    distance: float
    point_a: np.ndarray
    point_b: np.ndarray
    triangle_a: int
    triangle_b: int
    method: str


@dataclass(slots=True)
class Interference:
    collides: bool
    volume: float | None
    method: str
    note: str = ""


def aabb_distance(a: trimesh.Trimesh, b: trimesh.Trimesh) -> float:
    amin, amax = np.asarray(a.bounds)
    bmin, bmax = np.asarray(b.bounds)
    delta = np.maximum(0.0, np.maximum(amin - bmax, bmin - amax))
    return float(np.linalg.norm(delta))


def aabb_overlaps(a: trimesh.Trimesh, b: trimesh.Trimesh, padding: float = 0.0) -> bool:
    amin, amax = np.asarray(a.bounds)
    bmin, bmax = np.asarray(b.bounds)
    return bool(np.all(amax + padding >= bmin) and np.all(bmax + padding >= amin))


def candidate_pairs(bodies: list[BodyRecord], max_aabb_distance: float | None = None):
    for i, a in enumerate(bodies):
        for b in bodies[i + 1 :]:
            bound = aabb_distance(a.mesh, b.mesh)
            if max_aabb_distance is None or bound <= max_aabb_distance:
                yield a, b, bound


def _as_mesh(value) -> trimesh.Trimesh | None:
    if isinstance(value, trimesh.Trimesh):
        return value
    if isinstance(value, trimesh.Scene):
        dumped = value.dump(concatenate=True)
        return dumped if isinstance(dumped, trimesh.Trimesh) else None
    if isinstance(value, (list, tuple)):
        meshes = [item for item in value if isinstance(item, trimesh.Trimesh)]
        return trimesh.util.concatenate(meshes) if meshes else None
    return None


def interference(a: trimesh.Trimesh, b: trimesh.Trimesh, volume_tolerance: float) -> Interference:
    if not aabb_overlaps(a, b):
        return Interference(False, 0.0, "aabb")
    if a.is_watertight and b.is_watertight:
        try:
            overlap = _as_mesh(trimesh.boolean.intersection([a, b], engine="manifold"))
            if overlap is None or overlap.is_empty:
                return Interference(False, 0.0, "manifold_boolean")
            volume = abs(float(overlap.volume)) if overlap.is_volume else 0.0
            return Interference(volume > volume_tolerance, volume, "manifold_boolean")
        except Exception as exc:
            boolean_note = str(exc)
    else:
        boolean_note = "one or both meshes are not watertight"

    try:
        manager = trimesh.collision.CollisionManager()
        manager.add_object("a", a)
        collided = bool(manager.in_collision_single(b))
        return Interference(collided, None, "python_fcl", boolean_note)
    except Exception:
        # Conservative fallback: interior vertices prove interference. A failure
        # to find one is inconclusive and is reported by the method/note.
        try:
            ainb = np.any(b.contains(_sample_points(a, 1000)))
            bina = np.any(a.contains(_sample_points(b, 1000)))
            collided = bool(ainb or bina)
        except Exception:
            collided = False
        return Interference(
            collided,
            None,
            "sampled_containment_fallback",
            f"Boolean/FCL unavailable ({boolean_note}); a negative result is not proof of no interference.",
        )


def _sample_points(mesh: trimesh.Trimesh, budget: int) -> np.ndarray:
    vertices = np.asarray(mesh.vertices)
    centroids = np.asarray(mesh.triangles_center)
    if len(vertices) > budget:
        vertices = vertices[np.linspace(0, len(vertices) - 1, budget, dtype=int)]
    remaining = max(0, budget - len(vertices))
    if remaining and len(centroids) > remaining:
        centroids = centroids[np.linspace(0, len(centroids) - 1, remaining, dtype=int)]
    elif remaining == 0:
        centroids = np.empty((0, 3))
    return np.vstack((vertices, centroids))


def _closest(mesh: trimesh.Trimesh, points: np.ndarray):
    try:
        return trimesh.proximity.closest_point(mesh, points)
    except Exception:
        # Vertex KD-tree fallback does not find face/edge interiors.
        from scipy.spatial import cKDTree

        distances, vertex_ids = cKDTree(np.asarray(mesh.vertices)).query(points, k=1)
        nearest = np.asarray(mesh.vertices)[vertex_ids]
        triangle_ids = np.full(len(points), -1, dtype=int)
        return nearest, distances, triangle_ids


def minimum_distance(a: trimesh.Trimesh, b: trimesh.Trimesh, samples: int = 2500) -> PairDistance:
    points_a = _sample_points(a, max(16, samples))
    points_b = _sample_points(b, max(16, samples))
    near_b, distance_ab, tri_b = _closest(b, points_a)
    near_a, distance_ba, tri_a = _closest(a, points_b)

    i = int(np.argmin(distance_ab))
    j = int(np.argmin(distance_ba))
    if float(distance_ab[i]) <= float(distance_ba[j]):
        source_triangle = _nearest_source_triangle(a, points_a[i])
        return PairDistance(
            float(distance_ab[i]), points_a[i], near_b[i], source_triangle, int(tri_b[i]), "bidirectional_surface",
        )
    source_triangle = _nearest_source_triangle(b, points_b[j])
    return PairDistance(
        float(distance_ba[j]), near_a[j], points_b[j], int(tri_a[j]), source_triangle, "bidirectional_surface",
    )


def _nearest_source_triangle(mesh: trimesh.Trimesh, point: np.ndarray) -> int:
    centers = np.asarray(mesh.triangles_center)
    if len(centers) == 0:
        return -1
    return int(np.argmin(np.einsum("ij,ij->i", centers - point, centers - point)))


def face_patch_map(mesh: trimesh.Trimesh) -> tuple[np.ndarray, list[dict]]:
    """Map triangles to planar connected facets; curved leftovers remain triangles."""
    mapping = np.full(len(mesh.faces), -1, dtype=int)
    patches: list[dict] = []
    next_id = 0
    for facet in mesh.facets:
        ids = np.asarray(facet, dtype=int)
        mapping[ids] = next_id
        area = float(np.sum(mesh.area_faces[ids]))
        centroid = np.average(mesh.triangles_center[ids], axis=0, weights=mesh.area_faces[ids])
        patches.append(
            {"patch_id": next_id, "area": area, "triangle_count": len(ids), "kind": "planar_facet", "centroid": centroid.tolist()}
        )
        next_id += 1
    for triangle_id in np.flatnonzero(mapping < 0):
        mapping[triangle_id] = next_id
        patches.append(
            {
                "patch_id": next_id,
                "area": float(mesh.area_faces[triangle_id]),
                "triangle_count": 1,
                "kind": "tessellation_triangle",
                "centroid": mesh.triangles_center[triangle_id].tolist(),
            }
        )
        next_id += 1
    return mapping, patches


def patch_for_triangle(mesh: trimesh.Trimesh, triangle_id: int) -> int | None:
    if triangle_id < 0 or triangle_id >= len(mesh.faces):
        return None
    mapping, _ = face_patch_map(mesh)
    return int(mapping[triangle_id])


def minimum_patch_distance(
    a: trimesh.Trimesh,
    b: trimesh.Trimesh,
    patch_a: int | None,
    patch_b: int | None,
    samples: int,
) -> PairDistance:
    """Minimum sampled distance restricted to reconstructed surface patches."""
    map_a, _ = face_patch_map(a)
    map_b, _ = face_patch_map(b)
    ids_a = np.flatnonzero(map_a == patch_a) if patch_a is not None else np.arange(len(a.faces))
    ids_b = np.flatnonzero(map_b == patch_b) if patch_b is not None else np.arange(len(b.faces))
    if len(ids_a) == 0:
        raise ValueError(f"Patch {patch_a} does not exist on body A")
    if len(ids_b) == 0:
        raise ValueError(f"Patch {patch_b} does not exist on body B")
    mesh_a = a.submesh([ids_a], append=True, repair=False)
    mesh_b = b.submesh([ids_b], append=True, repair=False)
    result = minimum_distance(mesh_a, mesh_b, samples)
    if 0 <= result.triangle_a < len(ids_a):
        result.triangle_a = int(ids_a[result.triangle_a])
    if 0 <= result.triangle_b < len(ids_b):
        result.triangle_b = int(ids_b[result.triangle_b])
    result.method = "restricted_patch_" + result.method
    return result


def matching_pairs(bodies: list[BodyRecord], rule: GapRule) -> Iterable[tuple[BodyRecord, BodyRecord]]:
    try:
        pattern_a = re.compile(rule.body_a, re.IGNORECASE)
        pattern_b = re.compile(rule.body_b, re.IGNORECASE)
    except re.error as exc:
        raise ValueError(f"Invalid gap-rule regular expression: {exc}") from exc
    seen: set[tuple[str, str]] = set()
    for a in bodies:
        if not pattern_a.search(a.name) and not pattern_a.search(a.source_model or ""):
            continue
        for b in bodies:
            if a.body_id == b.body_id:
                continue
            if not pattern_b.search(b.name) and not pattern_b.search(b.source_model or ""):
                continue
            key = tuple(sorted((a.body_id, b.body_id)))
            if key not in seen:
                seen.add(key)
                yield a, b


def finite_or_none(value: float | None) -> float | None:
    return value if value is not None and math.isfinite(value) else None
