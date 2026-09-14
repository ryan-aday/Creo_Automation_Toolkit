from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .models import AuditConfig
from .units import canonical_length_unit

LOGGER = logging.getLogger(__name__)


@dataclass(slots=True)
class CreoSnapshot:
    model_name: str
    model_type: str
    step_path: Path
    length_units: str
    components: list[dict[str, Any]] = field(default_factory=list)
    materials: dict[str, str | None] = field(default_factory=dict)
    cte_parameters: dict[str, float] = field(default_factory=dict)
    no_material_tags: set[str] = field(default_factory=set)
    feature_findings: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _flatten_bom(node: Any, parent_path: str = "root") -> list[dict[str, Any]]:
    """Flatten multiple CREOSON BOM response shapes without discarding paths."""
    out: list[dict[str, Any]] = []
    if isinstance(node, list):
        for item in node:
            out.extend(_flatten_bom(item, parent_path))
        return out
    if not isinstance(node, dict):
        return out

    record = {key: value for key, value in node.items() if key != "children"}
    if any(key in record for key in ("file", "name", "generic")):
        record.setdefault("seq_path", parent_path)
        out.append(record)
    for index, child in enumerate(node.get("children") or []):
        child_path = child.get("seq_path", f"{parent_path}.{index + 1}") if isinstance(child, dict) else parent_path
        out.extend(_flatten_bom(child, child_path))
    return out


def _file_name(record: dict[str, Any]) -> str | None:
    for key in ("file", "filename", "name", "generic"):
        value = record.get(key)
        if value:
            return str(value)
    return None


def _material_map(raw: Any) -> dict[str, str | None]:
    if isinstance(raw, str):
        return {}
    if isinstance(raw, dict):
        raw = [raw]
    result: dict[str, str | None] = {}
    for item in raw or []:
        if not isinstance(item, dict):
            continue
        file_ = _file_name(item)
        if not file_:
            continue
        material = item.get("material") or item.get("current_material") or item.get("value")
        result[file_.lower()] = str(material) if material not in (None, "") else None
    return result


def _parameter_value(parameters: list[dict[str, Any]], names: list[str]) -> float | None:
    wanted = {name.upper() for name in names}
    for parameter in parameters:
        if str(parameter.get("name", "")).upper() not in wanted:
            continue
        try:
            return float(parameter.get("value"))
        except (TypeError, ValueError):
            continue
    return None


class CreoBridge:
    """Read a Creo model through CREOSON and export one STEP audit snapshot."""

    def __init__(self, config: AuditConfig):
        self.config = config
        self.client: Any = None

    def connect(self) -> None:
        try:
            import creopyson
        except ImportError as exc:  # pragma: no cover - environment dependent
            raise RuntimeError("creopyson is not installed; run `pip install -r requirements.txt`.") from exc
        self.client = creopyson.Client(self.config.creoson_host, self.config.creoson_port)
        self.client.connect()
        if self.config.creo_version is not None:
            self.client.creo_set_creo_version(self.config.creo_version)

    def close(self) -> None:
        if self.client is not None:
            try:
                self.client.disconnect()
            except Exception:  # pragma: no cover - best-effort cleanup
                LOGGER.debug("CREOSON disconnect failed", exc_info=True)

    def capture(self) -> CreoSnapshot:
        if self.client is None:
            self.connect()

        model = self.config.model
        if model.suffix.lower() not in {".prt", ".asm"}:
            raise ValueError("model_path must reference a Creo .prt or .asm file")
        if not model.parent.exists():
            raise FileNotFoundError(f"Creo model directory does not exist: {model.parent}")

        output_dir = Path(self.config.export_directory).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        self.client.file_open(
            file_=model.name,
            dirname=str(model.parent),
            display=True,
            activate=True,
            regen_force=False,
        )
        model_name = str(self.client.file_get_active().get("file", model.name))
        warnings: list[str] = []

        regeneration_error = None
        try:
            self.client.file_regenerate(file_=model_name, display=False)
        except Exception as exc:  # CREOSON uses Warning-derived errors in some releases
            regeneration_error = str(exc)
            warnings.append(f"Creo regeneration reported: {exc}")

        units_raw = self.client.file_get_length_units(file_=model_name)
        if isinstance(units_raw, dict):
            units_raw = units_raw.get("units") or units_raw.get("unit")
        units = canonical_length_unit(str(units_raw) if units_raw else None)

        components: list[dict[str, Any]] = []
        if model.suffix.lower() == ".asm":
            try:
                bom = self.client.bom_get_paths(
                    file_=model_name,
                    paths=True,
                    skeletons=True,
                    top_level=False,
                    get_transforms=True,
                    exclude_inactive=False,
                )
                components = _flatten_bom(bom.get("children", bom))
            except Exception as exc:
                warnings.append(f"BOM metadata could not be read: {exc}")
        else:
            components = [{"file": model_name, "seq_path": "root"}]

        session_models = self._session_model_names(model_name, components)
        feature_findings: list[dict[str, Any]] = []
        for session_model in session_models:
            feature_findings.extend(self._read_feature_health(session_model))
        if regeneration_error:
            feature_findings.insert(
                0,
                {
                    "status": "REGENERATION_ERROR",
                    "name": model_name,
                    "type": "MODEL",
                    "file": model_name,
                    "error": regeneration_error,
                },
            )

        session_parts = [name for name in session_models if name.lower().endswith(".prt")]
        materials = self._read_materials(session_parts)
        if session_parts and not materials:
            warnings.append("No part material assignments could be returned by CREOSON.")
        cte_parameters = self._read_cte_parameters(session_parts)
        no_material_tags = self._read_no_material_tags(session_parts)

        step_name = f"{model.stem}_geometry_audit.step"
        exported = self.client.interface_export_file(
            "STEP",
            file_=model_name,
            filename=step_name,
            dirname=str(output_dir),
            geom_flags="solids",
            advanced=False,
        )
        step_path = self._resolve_export(exported, output_dir, step_name)
        if not step_path.exists():
            raise FileNotFoundError(
                f"CREOSON reported a STEP export, but the file was not found at {step_path}. "
                "Confirm that CREOSON and this Python process see the same filesystem."
            )

        return CreoSnapshot(
            model_name=model_name,
            model_type=model.suffix.lower().lstrip("."),
            step_path=step_path,
            length_units=units,
            components=components,
            materials=materials,
            cte_parameters=cte_parameters,
            no_material_tags=no_material_tags,
            feature_findings=feature_findings,
            warnings=warnings,
        )

    def _read_feature_health(self, model_name: str) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        try:
            features = self.client.feature_list(
                file_=model_name,
                paths=True,
                inc_unnamed=True,
                no_datum=False,
                no_comp=False,
            )
        except Exception as exc:
            return [{"status": "QUERY_ERROR", "name": model_name, "type": "MODEL", "error": str(exc)}]
        for feature in features or []:
            status = str(feature.get("status", "")).upper()
            if status in {"UNREGENERATED", "INACTIVE"}:
                results.append(feature)
        return results

    def _session_model_names(self, model_name: str, components: list[dict[str, Any]]) -> list[str]:
        names = {_file_name(item) for item in components}
        names.discard(None)
        try:
            for item in self.client.file_list(file_=["*.prt", "*.asm"]) or []:
                if isinstance(item, str):
                    names.add(item)
                elif isinstance(item, dict) and _file_name(item):
                    names.add(_file_name(item))
        except Exception:
            LOGGER.debug("Unable to enumerate session parts", exc_info=True)
        names.add(model_name)
        return sorted(
            str(name) for name in names
            if str(name).lower().endswith((".prt", ".asm"))
        )

    def _read_materials(self, part_names: list[str]) -> dict[str, str | None]:
        try:
            wildcard = self.client.file_get_cur_material_wildcard(
                file_="*.prt", include_non_matching_parts=True
            )
            result = _material_map(wildcard)
        except Exception:
            result = {}
        for part in part_names:
            key = part.lower()
            if key in result:
                continue
            try:
                value = self.client.file_get_cur_material(file_=part)
                result[key] = str(value) if value not in (None, "") else None
            except Exception:
                result[key] = None
        return result

    def _read_cte_parameters(self, part_names: list[str]) -> dict[str, float]:
        result: dict[str, float] = {}
        for part in part_names:
            try:
                parameters = self.client.parameter_list(
                    name=self.config.cte_parameter_names,
                    file_=part,
                    encoded=False,
                )
                value = _parameter_value(parameters or [], self.config.cte_parameter_names)
                if value is not None:
                    result[part.lower()] = value
            except Exception:
                LOGGER.debug("Unable to read CTE parameter for %s", part, exc_info=True)
        return result

    def _read_no_material_tags(self, part_names: list[str]) -> set[str]:
        tagged: set[str] = set()
        tag_name = self.config.material_tag_parameter.strip()
        if not tag_name:
            return tagged
        for part in part_names:
            try:
                parameters = self.client.parameter_list(name=tag_name, file_=part, encoded=False)
            except Exception:
                continue
            for parameter in parameters or []:
                value = parameter.get("value")
                if value is True or str(value).strip().upper() in {"1", "TRUE", "YES", "Y", "NO_MATERIAL"}:
                    tagged.add(part.lower())
        return tagged

    @staticmethod
    def _resolve_export(raw: Any, output_dir: Path, fallback_name: str) -> Path:
        if isinstance(raw, dict):
            filename = raw.get("filename") or fallback_name
            dirname = raw.get("dirname") or str(output_dir)
            path = Path(filename)
            return path if path.is_absolute() else Path(dirname) / path
        return output_dir / fallback_name
