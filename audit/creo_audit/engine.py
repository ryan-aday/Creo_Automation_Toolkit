from __future__ import annotations

import logging
from dataclasses import asdict
from pathlib import Path

from .checks import (
    candidate_pairs,
    face_patch_map,
    interference,
    matching_pairs,
    minimum_distance,
    minimum_patch_distance,
    patch_for_triangle,
)
from .creo_bridge import CreoBridge, CreoSnapshot
from .mesh_io import load_step_bodies
from .models import AuditConfig, AuditResult, BodyRecord, Finding
from .thermal import thermally_expand
from .units import cubed_unit, squared_unit

LOGGER = logging.getLogger(__name__)


class AuditEngine:
    def __init__(self, config: AuditConfig):
        self.config = config

    def run(self) -> AuditResult:
        bridge = CreoBridge(self.config)
        try:
            snapshot = bridge.capture()
        finally:
            bridge.close()
        bodies, mesh_warnings = load_step_bodies(
            snapshot.step_path, snapshot.length_units, snapshot.components
        )
        return self.analyze_snapshot(snapshot, bodies, mesh_warnings)

    def analyze_snapshot(
        self,
        snapshot: CreoSnapshot,
        bodies: list[BodyRecord],
        extra_warnings: list[str] | None = None,
    ) -> AuditResult:
        self._attach_metadata(snapshot, bodies)
        findings: list[Finding] = []
        warnings = [*snapshot.warnings, *(extra_warnings or [])]
        colliding_ids: set[str] = set()

        for body in bodies:
            if not body.mesh.is_watertight:
                findings.append(
                    Finding(
                        "mesh_health",
                        "warning",
                        "Tessellation is not watertight; interference volume and containment may be inconclusive.",
                        body_a=body.name,
                    )
                )

        if self.config.run_interference:
            for a, b, _ in candidate_pairs(bodies, 0.0):
                hit = interference(a.mesh, b.mesh, self.config.interference_volume_tolerance)
                if hit.collides:
                    colliding_ids.update((a.body_id, b.body_id))
                    findings.append(
                        Finding(
                            "interference",
                            "error",
                            "Bodies have positive-volume interference or a confirmed triangle collision.",
                            body_a=a.name,
                            body_b=b.name,
                            value=hit.volume,
                            units=cubed_unit(snapshot.length_units) if hit.volume is not None else "",
                            details={"method": hit.method, "note": hit.note},
                        )
                    )
                elif hit.note:
                    warnings.append(f"{a.name} / {b.name}: {hit.note}")

        if self.config.run_global_gap and self.config.minimum_global_gap > 0:
            for a, b, bound in candidate_pairs(bodies, self.config.minimum_global_gap):
                distance = minimum_distance(a.mesh, b.mesh, self.config.distance_samples)
                if distance.distance < self.config.minimum_global_gap:
                    findings.append(
                        self._gap_finding(
                            "minimum_gap",
                            a,
                            b,
                            distance,
                            self.config.minimum_global_gap,
                            None,
                            0.0,
                            snapshot.length_units,
                            "Nominal global clearance is below the required minimum.",
                        )
                    )

        findings.extend(self._check_gap_rules(bodies, snapshot.length_units, state="nominal"))

        if self.config.run_materials:
            findings.extend(self._check_materials(snapshot, bodies))

        if self.config.run_feature_health:
            findings.extend(self._check_features(snapshot))

        if self.config.run_small_faces and self.config.minimum_face_area > 0:
            findings.extend(self._check_small_faces(bodies, snapshot.length_units))

        if self.config.operating_temperature is not None:
            thermal_findings, thermal_warnings = self._check_thermal_gaps(
                bodies, snapshot.length_units
            )
            findings.extend(thermal_findings)
            warnings.extend(thermal_warnings)

        body_summary = [
            {
                "body_id": body.body_id,
                "name": body.name,
                "source_model": body.source_model,
                "material": body.material,
                "cte_per_k": body.cte_per_k,
                "watertight": bool(body.mesh.is_watertight),
                "vertices": int(len(body.mesh.vertices)),
                "triangles": int(len(body.mesh.faces)),
            }
            for body in bodies
        ]
        metadata = {
            "step_path": str(snapshot.step_path),
            "components": snapshot.components,
            "body_summary": body_summary,
            "colliding_body_ids": sorted(colliding_ids),
            "analysis_basis": "tessellated STEP snapshot",
        }
        return AuditResult(
            model_name=snapshot.model_name,
            bodies=bodies,
            length_units=snapshot.length_units,
            findings=findings,
            warnings=list(dict.fromkeys(warnings)),
            metadata=metadata,
        )

    def _attach_metadata(self, snapshot: CreoSnapshot, bodies: list[BodyRecord]) -> None:
        for body in bodies:
            part_key = (body.source_model or "").lower()
            material = self._lookup_by_model(snapshot.materials, part_key, body.name)
            body.material = material
            raw_cte = self._lookup_by_model(snapshot.cte_parameters, part_key, body.name)
            if raw_cte is not None:
                body.cte_per_k = float(raw_cte) * self.config.cte_parameter_scale
            elif material:
                spec = self.config.material_library.get(material)
                if spec is None:
                    spec = next(
                        (value for key, value in self.config.material_library.items() if key.lower() == material.lower()),
                        None,
                    )
                body.cte_per_k = spec.cte_per_k if spec else None

    @staticmethod
    def _lookup_by_model(mapping: dict, part_key: str, body_name: str):
        if part_key in mapping:
            return mapping[part_key]
        part_stem = Path(part_key).stem
        probe = body_name.lower()
        candidates = []
        for key, value in mapping.items():
            stem = Path(key).stem.lower()
            if stem and (stem in probe or (part_stem and stem == part_stem)):
                candidates.append((len(stem), value))
        return max(candidates, default=(0, None))[1]

    def _check_materials(self, snapshot: CreoSnapshot, bodies: list[BodyRecord]) -> list[Finding]:
        default_names = {name.strip().upper() for name in self.config.default_material_names}
        findings: list[Finding] = []
        models = set(snapshot.materials) | set(snapshot.no_material_tags)
        if not models:
            return [
                Finding(
                    "material_assignment",
                    "warning",
                    "No part material records were available; the material check is inconclusive.",
                )
            ]
        for model in sorted(models):
            tagged = model in snapshot.no_material_tags
            material_value = snapshot.materials.get(model)
            material = (material_value or "").strip()
            if tagged or material.upper() in default_names:
                findings.append(
                    Finding(
                        "material_assignment",
                        "error",
                        "Part is tagged as having no material or uses a configured default/unset material name.",
                        body_a=model,
                        details={"material": material_value, "tagged_no_material": tagged},
                    )
                )
        return findings

    @staticmethod
    def _check_features(snapshot: CreoSnapshot) -> list[Finding]:
        findings: list[Finding] = []
        for feature in snapshot.feature_findings:
            status = str(feature.get("status", "UNKNOWN")).upper()
            severity = "error" if status in {"UNREGENERATED", "REGENERATION_ERROR", "QUERY_ERROR"} else "warning"
            findings.append(
                Finding(
                    "feature_health",
                    severity,
                    f"Creo feature/model status is {status}.",
                    body_a=str(feature.get("file") or snapshot.model_name),
                    details=feature,
                )
            )
        return findings

    def _check_small_faces(self, bodies: list[BodyRecord], units: str) -> list[Finding]:
        findings: list[Finding] = []
        for body in bodies:
            _, patches = face_patch_map(body.mesh)
            for patch in patches:
                if patch["area"] < self.config.minimum_face_area:
                    findings.append(
                        Finding(
                            "small_face",
                            "warning",
                            "A tessellated planar patch/triangle is below the configured area threshold.",
                            body_a=body.name,
                            value=patch["area"],
                            limit=self.config.minimum_face_area,
                            units=squared_unit(units),
                            details=patch,
                        )
                    )
        return findings

    def _check_gap_rules(self, bodies: list[BodyRecord], units: str, state: str) -> list[Finding]:
        findings: list[Finding] = []
        for rule_index, rule in enumerate(self.config.gap_rules, start=1):
            pairs = list(matching_pairs(bodies, rule))
            if not pairs:
                findings.append(
                    Finding(
                        f"gap_rule_{state}",
                        "warning",
                        f"Gap rule {rule_index} matched no body pairs.",
                        details={"rule": asdict(rule)},
                    )
                )
                continue
            for a, b in pairs:
                try:
                    distance = minimum_patch_distance(
                        a.mesh, b.mesh, rule.patch_a, rule.patch_b, self.config.distance_samples
                    )
                except ValueError as exc:
                    findings.append(
                        Finding(
                            f"gap_rule_{state}",
                            "error",
                            f"Gap rule {rule_index} could not be evaluated: {exc}",
                            body_a=a.name,
                            body_b=b.name,
                            details={"rule_index": rule_index, "patch_a": rule.patch_a, "patch_b": rule.patch_b},
                        )
                    )
                    continue
                effective = max(0.0, distance.distance - rule.stack_closure)
                allowed_max = rule.allowed_maximum
                passed = effective >= rule.allowed_minimum and (
                    allowed_max is None or effective <= allowed_max
                )
                message = (
                    f"{state.title()} clearance meets the rule."
                    if passed
                    else f"{state.title()} clearance is outside the allowed rule window."
                )
                findings.append(
                    self._gap_finding(
                        f"gap_rule_{state}",
                        a,
                        b,
                        distance,
                        rule.allowed_minimum,
                        allowed_max,
                        rule.stack_closure,
                        units,
                        message,
                        severity="pass" if passed else "error",
                        extra={
                            "rule_index": rule_index,
                            "description": rule.description,
                            "requested_patch_a": rule.patch_a,
                            "requested_patch_b": rule.patch_b,
                        },
                    )
                )
        return findings

    def _check_thermal_gaps(self, bodies: list[BodyRecord], units: str):
        delta = self.config.operating_temperature - self.config.reference_temperature
        expanded: dict[str, BodyRecord] = {}
        warnings: list[str] = []
        for body in bodies:
            if body.cte_per_k is None:
                warnings.append(
                    f"Thermal clearance skipped for {body.name}: no verified CTE was found."
                )
                continue
            try:
                expanded[body.body_id] = thermally_expand(body, delta, self.config.thermal_anchor)
            except ValueError as exc:
                warnings.append(str(exc))

        hot_bodies = [expanded[body.body_id] for body in bodies if body.body_id in expanded]
        if len(hot_bodies) < 2:
            return [], warnings
        findings = self._check_gap_rules(hot_bodies, units, state="thermal")
        for finding in findings:
            finding.details.update(
                {
                    "reference_temperature": self.config.reference_temperature,
                    "operating_temperature": self.config.operating_temperature,
                    "delta_temperature": delta,
                    "thermal_anchor": self.config.thermal_anchor,
                }
            )
        return findings, warnings

    @staticmethod
    def _gap_finding(
        check: str,
        a: BodyRecord,
        b: BodyRecord,
        distance,
        minimum: float,
        maximum: float | None,
        stack_closure: float,
        units: str,
        message: str,
        severity: str = "error",
        extra: dict | None = None,
    ) -> Finding:
        effective = max(0.0, distance.distance - stack_closure)
        details = {
            "raw_distance": distance.distance,
            "effective_distance": effective,
            "stack_closure": stack_closure,
            "allowed_maximum": maximum,
            "point_a": distance.point_a.tolist(),
            "point_b": distance.point_b.tolist(),
            "triangle_a": distance.triangle_a,
            "triangle_b": distance.triangle_b,
            "patch_a": patch_for_triangle(a.mesh, distance.triangle_a),
            "patch_b": patch_for_triangle(b.mesh, distance.triangle_b),
            "method": distance.method,
        }
        details.update(extra or {})
        return Finding(
            check,
            severity,
            message,
            body_a=a.name,
            body_b=b.name,
            value=effective,
            limit=minimum,
            units=units,
            details=details,
        )
