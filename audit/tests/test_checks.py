from pathlib import Path

import pytest
import trimesh

from creo_audit.checks import face_patch_map, interference, minimum_distance, minimum_patch_distance
from creo_audit.creo_bridge import CreoSnapshot
from creo_audit.engine import AuditEngine
from creo_audit.models import AuditConfig, BodyRecord, GapRule, MaterialSpec


def body(name: str, center, material=None) -> BodyRecord:
    mesh = trimesh.creation.box(extents=[1.0, 1.0, 1.0])
    mesh.apply_translation(center)
    return BodyRecord(name, name, name, name, f"{name}.prt", mesh, material=material)


def snapshot(**kwargs):
    values = dict(
        model_name="test.asm",
        model_type="asm",
        step_path=Path("test.step"),
        length_units="millimeters",
    )
    values.update(kwargs)
    return CreoSnapshot(**values)


def test_minimum_distance_between_boxes():
    a = body("a", [0, 0, 0])
    b = body("b", [2, 0, 0])
    result = minimum_distance(a.mesh, b.mesh, samples=1000)
    assert result.distance == pytest.approx(1.0, abs=1e-8)
    assert result.point_a[0] == pytest.approx(0.5)
    assert result.point_b[0] == pytest.approx(1.5)


def test_positive_volume_interference():
    a = body("a", [0, 0, 0])
    b = body("b", [0.75, 0, 0])
    result = interference(a.mesh, b.mesh, volume_tolerance=1e-10)
    assert result.collides
    assert result.volume == pytest.approx(0.25, rel=1e-5)


def test_cube_planar_patches_are_faces():
    a = body("a", [0, 0, 0])
    mapping, patches = face_patch_map(a.mesh)
    assert len(mapping) == 12
    assert len(patches) == 6
    assert all(patch["area"] == pytest.approx(1.0) for patch in patches)


def test_patch_restricted_distance():
    a = body("a", [0, 0, 0])
    b = body("b", [2, 0, 0])
    map_a, patches_a = face_patch_map(a.mesh)
    map_b, patches_b = face_patch_map(b.mesh)
    patch_a = max(patches_a, key=lambda p: p["centroid"][0])["patch_id"]
    patch_b = min(patches_b, key=lambda p: p["centroid"][0])["patch_id"]
    result = minimum_patch_distance(a.mesh, b.mesh, patch_a, patch_b, 1000)
    assert result.distance == pytest.approx(1.0)
    assert map_a[result.triangle_a] == patch_a
    assert map_b[result.triangle_b] == patch_b


def test_engine_gap_rule_and_material_tag():
    a = body("housing", [0, 0, 0])
    b = body("shaft", [2, 0, 0])
    config = AuditConfig(
        model_path="test.asm",
        run_interference=False,
        run_global_gap=False,
        run_small_faces=False,
        gap_rules=[GapRule("housing", "shaft", minimum=1.1)],
    )
    result = AuditEngine(config).analyze_snapshot(
        snapshot(
            materials={"housing.prt": "_NO_MATL_SET", "shaft.prt": "STEEL"},
            no_material_tags={"housing.prt"},
        ),
        [a, b],
    )
    assert any(item.check == "gap_rule_nominal" and item.severity == "error" for item in result.findings)
    assert any(item.check == "material_assignment" and item.severity == "error" for item in result.findings)


def test_thermal_gap_reduces_for_free_expansion():
    a = body("a", [0, 0, 0])
    b = body("b", [2, 0, 0])
    config = AuditConfig(
        model_path="test.asm",
        run_interference=False,
        run_global_gap=False,
        run_materials=False,
        run_feature_health=False,
        run_small_faces=False,
        reference_temperature=0.0,
        operating_temperature=100.0,
        material_library={"MAT": MaterialSpec("MAT", cte_per_k=1e-3)},
        gap_rules=[GapRule("a", "b", minimum=0.95)],
    )
    result = AuditEngine(config).analyze_snapshot(
        snapshot(materials={"a.prt": "MAT", "b.prt": "MAT"}),
        [a, b],
    )
    nominal = next(item for item in result.findings if item.check == "gap_rule_nominal")
    thermal = next(item for item in result.findings if item.check == "gap_rule_thermal")
    assert nominal.value == pytest.approx(1.0)
    assert thermal.value == pytest.approx(0.9)
    assert thermal.severity == "error"
