from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal


Severity = Literal["error", "warning", "info", "pass"]


@dataclass(slots=True)
class GapRule:
    """Required clearance between two body-name regular expressions."""

    body_a: str
    body_b: str
    minimum: float
    patch_a: int | None = None
    patch_b: int | None = None
    maximum: float | None = None
    target: float | None = None
    tolerance: float | None = None
    stack_closure: float = 0.0
    description: str = ""

    @property
    def allowed_minimum(self) -> float:
        if self.target is not None and self.tolerance is not None:
            return self.target - self.tolerance
        return self.minimum

    @property
    def allowed_maximum(self) -> float | None:
        if self.target is not None and self.tolerance is not None:
            return self.target + self.tolerance
        return self.maximum


@dataclass(slots=True)
class MaterialSpec:
    name: str
    cte_per_k: float | None = None
    source: str = ""


@dataclass(slots=True)
class AuditConfig:
    model_path: str
    creoson_host: str = "localhost"
    creoson_port: int = 9056
    creo_version: int | None = None
    export_directory: str = "audit_output"
    minimum_global_gap: float = 0.0
    minimum_face_area: float = 0.0
    interference_volume_tolerance: float = 1.0e-9
    contact_tolerance: float = 1.0e-7
    distance_samples: int = 2500
    default_material_names: list[str] = field(
        default_factory=lambda: ["", "DEFAULT", "_NO_MATL_SET", "NO_MATERIAL"]
    )
    material_tag_parameter: str = "_NO_MATL_SET"
    cte_parameter_names: list[str] = field(
        default_factory=lambda: [
            "CTE",
            "ALPHA",
            "THERMAL_EXPANSION_COEFFICIENT",
            "COEFF_THERMAL_EXPANSION",
        ]
    )
    cte_parameter_scale: float = 1.0e-6
    reference_temperature: float = 20.0
    operating_temperature: float | None = None
    thermal_anchor: Literal["centroid", "global_origin"] = "centroid"
    material_library: dict[str, MaterialSpec] = field(default_factory=dict)
    gap_rules: list[GapRule] = field(default_factory=list)
    run_interference: bool = True
    run_global_gap: bool = True
    run_small_faces: bool = True
    run_materials: bool = True
    run_feature_health: bool = True

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "AuditConfig":
        data = dict(raw)
        data["gap_rules"] = [
            item if isinstance(item, GapRule) else GapRule(**item)
            for item in data.get("gap_rules", [])
        ]
        materials: dict[str, MaterialSpec] = {}
        for key, value in data.get("material_library", {}).items():
            if isinstance(value, MaterialSpec):
                materials[key] = value
            else:
                item = dict(value)
                item.setdefault("name", key)
                materials[key] = MaterialSpec(**item)
        data["material_library"] = materials
        return cls(**data)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @property
    def model(self) -> Path:
        return Path(self.model_path).expanduser().resolve()


@dataclass(slots=True)
class BodyRecord:
    body_id: str
    name: str
    geometry_name: str
    node_name: str
    source_model: str | None
    mesh: Any
    material: str | None = None
    cte_per_k: float | None = None


@dataclass(slots=True)
class Finding:
    check: str
    severity: Severity
    message: str
    body_a: str | None = None
    body_b: str | None = None
    value: float | None = None
    limit: float | None = None
    units: str = ""
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class AuditResult:
    model_name: str
    length_units: str
    bodies: list[BodyRecord]
    findings: list[Finding]
    metadata: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def findings_dicts(self) -> list[dict[str, Any]]:
        return [finding.to_dict() for finding in self.findings]

    def summary(self) -> dict[str, int]:
        counts = {key: 0 for key in ("error", "warning", "info", "pass")}
        for finding in self.findings:
            counts[finding.severity] += 1
        return counts
