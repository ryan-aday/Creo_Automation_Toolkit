#!/usr/bin/env python3
"""
Creo Assembly Mass / Ownership / Material Report
================================================

Builds a hierarchy-preserving Excel BOM from a live Creo Parametric session
through creopyson / CREOSON.

Design goals
------------
* Preserve occurrence hierarchy and component-path indices.
* Collapse only duplicate components that share the SAME immediate parent
  occurrence. This prevents identical hardware used by different subassemblies
  from disappearing into a single global row.
* Also provide a separate global unique-model rollup.
* Read part mass/material/metadata once per unique model and cache it.
* Derive assembly masses recursively from their children instead of trusting
  assembly mass-property calls.
* Read Windchill/Creo ownership metadata from configurable model parameters.
* Optionally export and embed model thumbnails.
* Produce warning and summary sheets plus mass charts.
* Stay console-first, with progress bars and explicit warnings.

The script does not require a paid PTC Toolkit license; it assumes Creo,
CREOSON, and creopyson are already configured and available.
"""

from __future__ import annotations

import argparse
import collections
import dataclasses
import json
import math
import re
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

try:
    import creopyson
except ImportError as exc:
    raise SystemExit(
        "creopyson is not installed. Run: pip install -r requirements.txt"
    ) from exc

try:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit(
        "openpyxl is not installed. Run: pip install -r requirements.txt"
    ) from exc

try:
    from tqdm import tqdm
except ImportError as exc:
    raise SystemExit(
        "tqdm is not installed. Run: pip install -r requirements.txt"
    ) from exc


APP_NAME = "Creo Assembly Mass Report"
APP_VERSION = "2.0.0"

DEFAULT_OWNER_PARAMETERS = [
    "OWNER",
    "DESIGNER",
    "DRAWN_BY",
    "PTC_WM_CREATED_BY",
]
DEFAULT_CREATOR_PARAMETERS = [
    "PTC_WM_CREATED_BY",
    "CREATED_BY",
    "CREATOR",
]
DEFAULT_MODIFIER_PARAMETERS = [
    "PTC_WM_MODIFIED_BY",
    "MODIFIED_BY",
    "LAST_MODIFIED_BY",
]
DEFAULT_CREATED_ON_PARAMETERS = [
    "PTC_WM_CREATED_ON",
    "CREATED_ON",
]
DEFAULT_MODIFIED_ON_PARAMETERS = [
    "PTC_WM_MODIFIED_ON",
    "MODIFIED_ON",
]
DEFAULT_THUMBNAIL_VIEWS = [
    "ISO",
    "ISOMETRIC",
    "TRIMETRIC",
    "DEFAULT",
]

# Canonical factor to kilograms. Extend this dictionary if your Creo
# installation returns a site-specific unit string.
MASS_UNIT_TO_KG = {
    "kg": 1.0,
    "kilogram": 1.0,
    "kilograms": 1.0,
    "g": 1.0e-3,
    "gram": 1.0e-3,
    "grams": 1.0e-3,
    "mg": 1.0e-6,
    "milligram": 1.0e-6,
    "milligrams": 1.0e-6,
    "lb": 0.45359237,
    "lbs": 0.45359237,
    "lbm": 0.45359237,
    "pound": 0.45359237,
    "pounds": 0.45359237,
    "oz": 0.028349523125,
    "ounce": 0.028349523125,
    "ounces": 0.028349523125,
    "tonne": 1000.0,
    "tonnes": 1000.0,
    "metric_ton": 1000.0,
}

SEVERITY_ORDER = {"ERROR": 0, "WARNING": 1, "INFO": 2}


@dataclass
class WarningRecord:
    severity: str
    code: str
    model: str
    index: str = ""
    message: str = ""


@dataclass
class ModelMetadata:
    model: str
    model_type: str
    native_mass_unit: Optional[str] = None
    native_mass: float = math.inf
    material: Optional[str] = None
    owner: Optional[str] = None
    owner_parameter: Optional[str] = None
    created_by: Optional[str] = None
    created_by_parameter: Optional[str] = None
    modified_by: Optional[str] = None
    modified_by_parameter: Optional[str] = None
    created_on: Optional[str] = None
    modified_on: Optional[str] = None
    thumbnail_path: Optional[str] = None
    parameters: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


@dataclass
class Occurrence:
    path: str
    model: str
    depth: int
    parent_path: str
    parent_model: str
    hierarchy_models: list[str]
    raw: dict[str, Any] = field(default_factory=dict)
    children: list["Occurrence"] = field(default_factory=list)
    model_type: str = "PART"
    mass_kg: float = math.inf
    category: str = "OTHER"

    @property
    def is_assembly(self) -> bool:
        return self.model_type == "ASSEMBLY"

    @property
    def is_part(self) -> bool:
        return not self.is_assembly


@dataclass
class GroupRow:
    indices: list[str]
    model: str
    model_type: str
    immediate_assembly: str
    immediate_assembly_index: str
    assembly_path: str
    depth: int
    mass_kg_each: float
    qty: int
    total_mass_kg: float
    top_qty: int
    category: str
    metadata: ModelMetadata
    warnings: list[str] = field(default_factory=list)


@dataclass
class Settings:
    host: str = "localhost"
    port: int = 9056
    assembly: Optional[str] = None
    output: Path = Path("creo_assembly_mass_report.xlsx")
    thumbnail_dir: Path = Path("creo_report_thumbnails")
    thumbnails: bool = False
    thumbnail_width: float = 3.0
    thumbnail_height: float = 3.0
    thumbnail_dpi: int = 100
    thumbnail_views: list[str] = field(
        default_factory=lambda: list(DEFAULT_THUMBNAIL_VIEWS)
    )
    creo_version: Optional[int] = None
    report_unit: Optional[str] = None
    high_mass_warning: float = 50.0
    include_skeletons: bool = False
    owner_parameters: list[str] = field(
        default_factory=lambda: list(DEFAULT_OWNER_PARAMETERS)
    )
    creator_parameters: list[str] = field(
        default_factory=lambda: list(DEFAULT_CREATOR_PARAMETERS)
    )
    modifier_parameters: list[str] = field(
        default_factory=lambda: list(DEFAULT_MODIFIER_PARAMETERS)
    )
    created_on_parameters: list[str] = field(
        default_factory=lambda: list(DEFAULT_CREATED_ON_PARAMETERS)
    )
    modified_on_parameters: list[str] = field(
        default_factory=lambda: list(DEFAULT_MODIFIED_ON_PARAMETERS)
    )
    categories: collections.OrderedDict[str, list[str]] = field(
        default_factory=lambda: collections.OrderedDict(
            [
                ("FASTENER", ["NAS"]),
                ("HARNESS", ["HARNESS"]),
            ]
        )
    )


def console_info(message: str) -> None:
    tqdm.write(f"[INFO] {message}")


def console_warn(message: str) -> None:
    tqdm.write(f"[WARNING] {message}")


def console_error(message: str) -> None:
    tqdm.write(f"[ERROR] {message}")


def normalize_model_name(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def model_type_from_name(model: str, has_children: bool = False) -> str:
    lower = model.lower()
    if lower.endswith(".asm") or has_children:
        return "ASSEMBLY"
    return "PART"


def normalize_seq_path(raw_path: Any, fallback: str) -> str:
    """Normalize CREOSON seq_path values such as root.3.2 -> 0.3.2."""
    if raw_path is None:
        return fallback
    if isinstance(raw_path, (list, tuple)):
        tokens = [str(x) for x in raw_path]
        return ".".join(tokens) if tokens else fallback
    value = str(raw_path).strip()
    if not value:
        return fallback
    value = re.sub(r"^root(?=\.|$)", "0", value, flags=re.IGNORECASE)
    return value


def natural_path_key(path: str) -> tuple:
    parts = re.split(r"[./\\:_-]+", path)
    key = []
    for token in parts:
        if token.isdigit():
            key.append((0, int(token)))
        else:
            key.append((1, token.casefold()))
    return tuple(key)


def extract_child_nodes(node: Any) -> list[dict[str, Any]]:
    if isinstance(node, dict):
        children = node.get("children")
        if isinstance(children, list):
            return [x for x in children if isinstance(x, dict)]
        if isinstance(children, dict):
            return [children]
    return []


def extract_model_name(node: dict[str, Any]) -> str:
    for key in ("file", "filename", "model", "name", "instance"):
        value = node.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    # Some CREOSON versions can nest file information.
    for key in ("model_data", "component", "item"):
        value = node.get(key)
        if isinstance(value, dict):
            candidate = extract_model_name(value)
            if candidate:
                return candidate
    return ""


def parse_bom_tree(bom: dict[str, Any], top_model: str) -> Occurrence:
    """
    Convert the CREOSON BOM response into an occurrence tree.

    The parser deliberately tolerates small differences in JSON shape between
    CREOSON/creopyson versions.
    """
    root = Occurrence(
        path="0",
        model=top_model,
        depth=0,
        parent_path="",
        parent_model="",
        hierarchy_models=[top_model],
        raw=bom,
        model_type="ASSEMBLY",
    )

    children = extract_child_nodes(bom)
    if not children and isinstance(bom.get("children"), list):
        children = bom["children"]

    def walk(
        nodes: list[dict[str, Any]],
        parent: Occurrence,
    ) -> None:
        for ordinal, node in enumerate(nodes, start=1):
            model = extract_model_name(node)
            if not model:
                # Skip non-component bookkeeping nodes.
                continue
            fallback = f"{parent.path}.{ordinal}"
            seq_path = normalize_seq_path(
                node.get("seq_path", node.get("path")), fallback
            )
            child_nodes = extract_child_nodes(node)
            occ = Occurrence(
                path=seq_path,
                model=model,
                depth=parent.depth + 1,
                parent_path=parent.path,
                parent_model=parent.model,
                hierarchy_models=parent.hierarchy_models + [model],
                raw=node,
                model_type=model_type_from_name(model, bool(child_nodes)),
            )
            parent.children.append(occ)
            walk(child_nodes, occ)

    walk(children, root)
    return root


def iter_occurrences(root: Occurrence, include_root: bool = False) -> Iterable[Occurrence]:
    if include_root:
        yield root
    for child in root.children:
        yield child
        yield from iter_occurrences(child, include_root=False)


def find_first_parameter(
    param_map: dict[str, Any], priorities: Sequence[str]
) -> tuple[Optional[str], Optional[str]]:
    for name in priorities:
        key = name.upper()
        if key in param_map:
            value = param_map[key]
            if value is not None and str(value).strip():
                return str(value).strip(), name
    return None, None


def coerce_material(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, str):
        return value.strip() or None
    if isinstance(value, dict):
        for key in ("material", "name", "value"):
            if value.get(key):
                return str(value[key]).strip()
    if isinstance(value, list):
        if not value:
            return None
        return coerce_material(value[0])
    return str(value).strip() or None


def normalize_mass_unit(unit: Optional[str]) -> Optional[str]:
    if unit is None:
        return None
    s = str(unit).strip().lower()
    s = s.replace(" ", "_")
    aliases = {
        "kilogram_force_sec2/meter": "kg",  # defensive only; uncommon
        "pound_mass": "lbm",
        "lb_mass": "lbm",
        "kilograms": "kg",
        "grams": "g",
        "ounces": "oz",
    }
    return aliases.get(s, s)


def mass_to_kg(value: float, unit: Optional[str]) -> float:
    if not math.isfinite(value):
        return math.inf
    norm = normalize_mass_unit(unit)
    factor = MASS_UNIT_TO_KG.get(norm or "")
    if factor is None:
        return math.inf
    return value * factor


def kg_to_mass(value_kg: float, unit: str) -> float:
    if not math.isfinite(value_kg):
        return math.inf
    norm = normalize_mass_unit(unit)
    factor = MASS_UNIT_TO_KG.get(norm or "")
    if factor is None or factor == 0:
        return math.inf
    return value_kg / factor


def report_unit_supported(unit: Optional[str]) -> bool:
    return normalize_mass_unit(unit) in MASS_UNIT_TO_KG


def classify_model(model: str, categories: collections.OrderedDict[str, list[str]]) -> str:
    upper = model.upper()
    for category, patterns in categories.items():
        for pattern in patterns:
            p = pattern.strip().upper()
            if p and p in upper:
                return category
    return "OTHER"


class CreoAdapter:
    """Small compatibility layer around creopyson's Client methods."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self.client = creopyson.Client(settings.host, settings.port)

    def connect(self) -> None:
        self.client.connect()
        if self.settings.creo_version is not None:
            method = getattr(self.client, "creo_set_creo_version", None)
            if callable(method):
                method(int(self.settings.creo_version))

    def get_active(self) -> dict[str, Any]:
        result = self.client.file_get_active()
        return result or {}

    def open_model(self, model: str) -> None:
        self.client.file_open(model, display=True, activate=True)

    def display_model(self, model: str) -> None:
        method = getattr(self.client, "file_display", None)
        if callable(method):
            method(model, activate=True)
        else:
            self.client.file_open(model, display=True, activate=True)

    def get_bom(self, assembly: str) -> dict[str, Any]:
        return self.client.bom_get_paths(
            file_=assembly,
            paths=True,
            skeletons=self.settings.include_skeletons,
            top_level=False,
            get_transforms=False,
            exclude_inactive=True,
            get_simpreps=True,
        )

    def get_mass_units(self, model: str) -> Optional[str]:
        return self.client.file_get_mass_units(file_=model)

    def get_massprops(self, model: str) -> dict[str, Any]:
        result = self.client.file_massprops(file_=model)
        return result or {}

    def get_material(self, model: str) -> Optional[str]:
        return coerce_material(self.client.file_get_cur_material(file_=model))

    def get_parameters(self, model: str, names: Sequence[str]) -> dict[str, Any]:
        unique_names = list(dict.fromkeys(names))
        result = self.client.parameter_list(name=unique_names, file_=model)
        param_map: dict[str, Any] = {}
        for item in result or []:
            if not isinstance(item, dict):
                continue
            name = item.get("name")
            if name:
                param_map[str(name).upper()] = item.get("value")
        return param_map

    def list_views(self, model: str) -> list[str]:
        result = self.client.view_list(file_=model)
        return [str(x) for x in (result or [])]

    def activate_view(self, model: str, name: str) -> None:
        self.client.view_activate(name=name, file_=model)

    def export_jpeg(
        self,
        model: str,
        destination: Path,
        width: float,
        height: float,
        dpi: int,
    ) -> dict[str, Any]:
        return self.client.interface_export_image(
            "JPEG",
            file_=model,
            filename=str(destination),
            width=width,
            height=height,
            dpi=dpi,
            depth=24,
        )


def load_model_metadata(
    adapter: CreoAdapter,
    models: Sequence[str],
    settings: Settings,
    warnings: list[WarningRecord],
) -> dict[str, ModelMetadata]:
    cache: dict[str, ModelMetadata] = {}

    parameter_names = list(
        dict.fromkeys(
            settings.owner_parameters
            + settings.creator_parameters
            + settings.modifier_parameters
            + settings.created_on_parameters
            + settings.modified_on_parameters
        )
    )

    for model in tqdm(models, desc="Reading Creo model data", unit="model"):
        model_type = model_type_from_name(model)
        md = ModelMetadata(model=model, model_type=model_type)

        # Mass unit is useful for both part records and report-unit inference.
        try:
            md.native_mass_unit = adapter.get_mass_units(model)
        except Exception as exc:
            md.native_mass_unit = None
            md.warnings.append("UNIT_READ_FAILED")
            warnings.append(
                WarningRecord(
                    "WARNING",
                    "UNIT_READ_FAILED",
                    model,
                    message=f"Could not read mass units: {exc}",
                )
            )

        if model_type == "PART":
            try:
                props = adapter.get_massprops(model)
                raw_mass = props.get("mass", math.inf)
                md.native_mass = float(raw_mass)
            except Exception as exc:
                md.native_mass = math.inf
                md.warnings.append("MASS_READ_FAILED")
                warnings.append(
                    WarningRecord(
                        "ERROR",
                        "MASS_READ_FAILED",
                        model,
                        message=f"Could not read part mass: {exc}",
                    )
                )

            try:
                md.material = adapter.get_material(model)
            except Exception as exc:
                md.material = None
                md.warnings.append("MATERIAL_READ_FAILED")
                warnings.append(
                    WarningRecord(
                        "WARNING",
                        "MATERIAL_READ_FAILED",
                        model,
                        message=f"Could not read current material: {exc}",
                    )
                )

        try:
            md.parameters = adapter.get_parameters(model, parameter_names)
        except Exception as exc:
            md.parameters = {}
            md.warnings.append("PARAMETER_READ_FAILED")
            warnings.append(
                WarningRecord(
                    "WARNING",
                    "PARAMETER_READ_FAILED",
                    model,
                    message=f"Could not read metadata parameters: {exc}",
                )
            )

        md.owner, md.owner_parameter = find_first_parameter(
            md.parameters, settings.owner_parameters
        )
        md.created_by, md.created_by_parameter = find_first_parameter(
            md.parameters, settings.creator_parameters
        )
        md.modified_by, md.modified_by_parameter = find_first_parameter(
            md.parameters, settings.modifier_parameters
        )
        md.created_on, _ = find_first_parameter(
            md.parameters, settings.created_on_parameters
        )
        md.modified_on, _ = find_first_parameter(
            md.parameters, settings.modified_on_parameters
        )

        if not md.owner:
            md.warnings.append("OWNER_NOT_FOUND")
            warnings.append(
                WarningRecord(
                    "INFO",
                    "OWNER_NOT_FOUND",
                    model,
                    message="No configured owner parameter contained a value.",
                )
            )

        cache[model] = md

    return cache


def infer_report_unit(
    adapter: CreoAdapter,
    root_model: str,
    metadata: dict[str, ModelMetadata],
    requested: Optional[str],
    warnings: list[WarningRecord],
) -> str:
    if requested:
        norm = normalize_mass_unit(requested)
        if norm not in MASS_UNIT_TO_KG:
            raise ValueError(
                f"Unsupported --report-unit '{requested}'. "
                f"Supported: {', '.join(sorted(MASS_UNIT_TO_KG))}"
            )
        return norm

    try:
        root_unit = normalize_mass_unit(adapter.get_mass_units(root_model))
        if root_unit in MASS_UNIT_TO_KG:
            return root_unit
    except Exception:
        pass

    for md in metadata.values():
        norm = normalize_mass_unit(md.native_mass_unit)
        if norm in MASS_UNIT_TO_KG:
            return norm

    warnings.append(
        WarningRecord(
            "ERROR",
            "REPORT_UNIT_FALLBACK",
            root_model,
            message=(
                "Could not infer a recognized mass unit from Creo. Falling back "
                "to kg; models with unrecognized units will remain Inf."
            ),
        )
    )
    return "kg"


def validate_part_metadata(
    metadata: dict[str, ModelMetadata],
    settings: Settings,
    report_unit: str,
    warnings: list[WarningRecord],
) -> None:
    for model, md in metadata.items():
        if md.model_type != "PART":
            continue

        if not report_unit_supported(md.native_mass_unit):
            md.warnings.append("UNIT_UNSUPPORTED")
            warnings.append(
                WarningRecord(
                    "ERROR",
                    "UNIT_UNSUPPORTED",
                    model,
                    message=(
                        f"Mass unit '{md.native_mass_unit}' is unsupported; "
                        "mass cannot be safely combined with other models."
                    ),
                )
            )

        mass_kg = mass_to_kg(md.native_mass, md.native_mass_unit)
        if not math.isfinite(md.native_mass) or not math.isfinite(mass_kg):
            if "MASS_READ_FAILED" not in md.warnings:
                md.warnings.append("MASS_INVALID")
                warnings.append(
                    WarningRecord(
                        "ERROR",
                        "MASS_INVALID",
                        model,
                        message="Mass is not finite after unit conversion.",
                    )
                )
        else:
            report_mass = kg_to_mass(mass_kg, report_unit)
            if report_mass > settings.high_mass_warning:
                md.warnings.append("MASS_HIGH")
                warnings.append(
                    WarningRecord(
                        "WARNING",
                        "MASS_HIGH",
                        model,
                        message=(
                            f"Part mass {report_mass:.6g} {report_unit} exceeds "
                            f"threshold {settings.high_mass_warning:g} {report_unit}."
                        ),
                    )
                )
            if report_mass <= 0:
                md.warnings.append("MASS_NONPOSITIVE")
                warnings.append(
                    WarningRecord(
                        "ERROR",
                        "MASS_NONPOSITIVE",
                        model,
                        message=f"Part mass is {report_mass:.6g} {report_unit}.",
                    )
                )

        mat = (md.material or "").upper()
        if "_NO_MATL_ASSIGNED" in mat or "DEFAULT" in mat or not mat.strip():
            md.warnings.append("MATERIAL_DEFAULT_OR_MISSING")
            warnings.append(
                WarningRecord(
                    "WARNING",
                    "MATERIAL_DEFAULT_OR_MISSING",
                    model,
                    message=f"Material is '{md.material or '<missing>'}'.",
                )
            )


def assign_occurrence_properties(
    root: Occurrence,
    metadata: dict[str, ModelMetadata],
    categories: collections.OrderedDict[str, list[str]],
) -> None:
    for occ in iter_occurrences(root, include_root=False):
        occ.category = classify_model(occ.model, categories)
        if occ.is_part:
            md = metadata[occ.model]
            occ.mass_kg = mass_to_kg(md.native_mass, md.native_mass_unit)

    def derive(occ: Occurrence) -> float:
        if occ.is_part:
            return occ.mass_kg
        total = 0.0
        for child in occ.children:
            child_mass = derive(child)
            if not math.isfinite(child_mass):
                occ.mass_kg = math.inf
                return math.inf
            total += child_mass
        occ.mass_kg = total
        return total

    derive(root)


def mass_excluding_category_kg(occ: Occurrence, excluded: str) -> float:
    if occ.is_part:
        if occ.category.upper() == excluded.upper():
            return 0.0
        return occ.mass_kg
    total = 0.0
    for child in occ.children:
        value = mass_excluding_category_kg(child, excluded)
        if not math.isfinite(value):
            return math.inf
        total += value
    return total


def add_occurrence_warning_indices(
    root: Occurrence,
    metadata: dict[str, ModelMetadata],
    warnings: list[WarningRecord],
) -> None:
    # Existing model-level warnings are useful, but this adds occurrence paths
    # for mass propagation problems.
    for occ in iter_occurrences(root):
        if not math.isfinite(occ.mass_kg):
            warnings.append(
                WarningRecord(
                    "ERROR",
                    "OCCURRENCE_MASS_INVALID",
                    occ.model,
                    index=occ.path,
                    message=(
                        "Occurrence mass is non-finite. For assemblies this means "
                        "one or more descendant part masses/units could not be resolved."
                    ),
                )
            )


def top_model_quantities(root: Occurrence) -> collections.Counter[str]:
    return collections.Counter(occ.model for occ in iter_occurrences(root))


def group_hierarchy_rows(
    root: Occurrence,
    metadata: dict[str, ModelMetadata],
) -> list[GroupRow]:
    """
    Collapse only rows that have the same model AND same immediate parent
    occurrence path.
    """
    top_qty = top_model_quantities(root)
    groups: dict[tuple[str, str], list[Occurrence]] = collections.defaultdict(list)
    for occ in iter_occurrences(root):
        groups[(occ.parent_path, occ.model)].append(occ)

    rows: list[GroupRow] = []
    for (parent_path, model), occurrences in groups.items():
        occurrences.sort(key=lambda o: natural_path_key(o.path))
        first = occurrences[0]
        masses = [o.mass_kg for o in occurrences]
        total_mass = (
            sum(masses) if all(math.isfinite(v) for v in masses) else math.inf
        )
        equal_mass = (
            all(math.isclose(masses[0], m, rel_tol=1e-9, abs_tol=1e-12) for m in masses[1:])
            if masses and all(math.isfinite(v) for v in masses)
            else False
        )
        each_mass = masses[0] if len(masses) == 1 or equal_mass else math.inf

        row_warnings = list(metadata[model].warnings)
        if len(masses) > 1 and not equal_mass:
            row_warnings.append("OCCURRENCE_MASS_VARIES")

        assembly_path = " > ".join(first.hierarchy_models[:-1])

        rows.append(
            GroupRow(
                indices=[o.path for o in occurrences],
                model=model,
                model_type=first.model_type,
                immediate_assembly=first.parent_model,
                immediate_assembly_index=parent_path,
                assembly_path=assembly_path,
                depth=first.depth,
                mass_kg_each=each_mass,
                qty=len(occurrences),
                total_mass_kg=total_mass,
                top_qty=top_qty[model],
                category=first.category,
                metadata=metadata[model],
                warnings=row_warnings,
            )
        )

    rows.sort(
        key=lambda r: (
            natural_path_key(r.immediate_assembly_index or "0"),
            min(natural_path_key(i) for i in r.indices),
            r.model.casefold(),
        )
    )
    return rows


def group_unique_rows(
    root: Occurrence,
    metadata: dict[str, ModelMetadata],
) -> list[GroupRow]:
    groups: dict[str, list[Occurrence]] = collections.defaultdict(list)
    for occ in iter_occurrences(root):
        groups[occ.model].append(occ)

    rows: list[GroupRow] = []
    for model, occurrences in groups.items():
        occurrences.sort(key=lambda o: natural_path_key(o.path))
        first = occurrences[0]
        masses = [o.mass_kg for o in occurrences]
        finite = all(math.isfinite(v) for v in masses)
        equal_mass = (
            finite
            and all(
                math.isclose(masses[0], m, rel_tol=1e-9, abs_tol=1e-12)
                for m in masses[1:]
            )
        )
        each = masses[0] if len(masses) == 1 or equal_mass else math.inf
        total = sum(masses) if finite else math.inf
        parents = sorted(
            {o.parent_model for o in occurrences if o.parent_model},
            key=str.casefold,
        )
        parent_paths = sorted(
            {o.parent_path for o in occurrences if o.parent_path},
            key=natural_path_key,
        )
        row_warnings = list(metadata[model].warnings)
        if len(masses) > 1 and not equal_mass:
            row_warnings.append("OCCURRENCE_MASS_VARIES")

        rows.append(
            GroupRow(
                indices=[o.path for o in occurrences],
                model=model,
                model_type=first.model_type,
                immediate_assembly="; ".join(parents),
                immediate_assembly_index="; ".join(parent_paths),
                assembly_path="; ".join(
                    sorted(
                        {" > ".join(o.hierarchy_models[:-1]) for o in occurrences},
                        key=str.casefold,
                    )
                ),
                depth=min(o.depth for o in occurrences),
                mass_kg_each=each,
                qty=len(occurrences),
                total_mass_kg=total,
                top_qty=len(occurrences),
                category=first.category,
                metadata=metadata[model],
                warnings=row_warnings,
            )
        )

    rows.sort(key=lambda r: (r.model_type != "ASSEMBLY", r.model.casefold()))
    return rows


def sanitize_filename(name: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "_", name)
    return value.rstrip(". ") or "model"


def choose_view(available: Sequence[str], preferred: Sequence[str]) -> Optional[str]:
    if not available:
        return None
    by_upper = {v.upper(): v for v in available}
    for wanted in preferred:
        w = wanted.upper()
        if w in by_upper:
            return by_upper[w]
    for wanted in preferred:
        w = wanted.upper()
        for actual in available:
            if w in actual.upper():
                return actual
    return available[0]


def generate_thumbnails(
    adapter: CreoAdapter,
    models: Sequence[str],
    metadata: dict[str, ModelMetadata],
    settings: Settings,
    warnings: list[WarningRecord],
    restore_model: str,
) -> None:
    if not settings.thumbnails:
        return

    settings.thumbnail_dir.mkdir(parents=True, exist_ok=True)
    console_info(f"Thumbnail directory: {settings.thumbnail_dir.resolve()}")

    for model in tqdm(models, desc="Exporting thumbnails", unit="model"):
        target = settings.thumbnail_dir / f"{sanitize_filename(model)}.jpg"
        try:
            adapter.display_model(model)
            available = adapter.list_views(model)
            selected = choose_view(available, settings.thumbnail_views)
            if selected:
                adapter.activate_view(model, selected)
            result = adapter.export_jpeg(
                model=model,
                destination=target.resolve(),
                width=settings.thumbnail_width,
                height=settings.thumbnail_height,
                dpi=settings.thumbnail_dpi,
            )

            # CREOSON may return an alternate path/name. Prefer the requested
            # path if it exists, otherwise reconstruct from response.
            final_path = target.resolve()
            if not final_path.exists() and isinstance(result, dict):
                dirname = result.get("dirname")
                filename = result.get("filename")
                if filename:
                    candidate = Path(dirname or "") / str(filename)
                    if candidate.exists():
                        final_path = candidate.resolve()

            if not final_path.exists():
                raise FileNotFoundError(
                    "CREOSON reported image export success but the JPEG "
                    f"was not found at {target.resolve()}"
                )
            metadata[model].thumbnail_path = str(final_path)
        except Exception as exc:
            metadata[model].warnings.append("THUMBNAIL_FAILED")
            warnings.append(
                WarningRecord(
                    "WARNING",
                    "THUMBNAIL_FAILED",
                    model,
                    message=f"Could not export thumbnail: {exc}",
                )
            )
            console_warn(f"Thumbnail failed for {model}: {exc}")

    try:
        adapter.display_model(restore_model)
    except Exception as exc:
        console_warn(f"Could not restore top assembly display: {exc}")


def depth_fill(depth: int, core: bool = False) -> PatternFill:
    if core:
        return PatternFill("solid", fgColor="2F75B5")
    # Blue-grey gradient. Clamp to avoid becoming pure white.
    palette = [
        "5B9BD5",
        "8FBCE6",
        "B4D3ED",
        "D5E5F4",
        "E8F0F8",
    ]
    idx = max(0, min(depth - 1, len(palette) - 1))
    return PatternFill("solid", fgColor=palette[idx])


CATEGORY_FONT_COLORS = [
    "C65911",
    "7030A0",
    "548235",
    "BF9000",
    "2F5597",
    "A61C00",
    "008C95",
]


def format_mass(value: float, report_unit: str) -> Optional[float]:
    if not math.isfinite(value):
        return None
    return kg_to_mass(value, report_unit)


def warnings_text(codes: Sequence[str]) -> str:
    return "; ".join(sorted(set(codes)))


def write_bom_sheet(
    wb: Workbook,
    sheet_name: str,
    rows: Sequence[GroupRow],
    report_unit: str,
    settings: Settings,
    unique_sheet: bool = False,
) -> None:
    ws = wb.create_sheet(sheet_name)
    headers = [
        "Index",
        "Name",
        "Mass Units",
        "Mass",
        "Qty",
        "Total Mass",
        "Type",
        "Immediate Assembly",
        "Immediate Assembly Index",
        "Assembly Path",
        "Depth",
        "Qty in Top Assembly",
        "Native Mass Units",
        "Material",
        "Category",
        "Owner",
        "Owner Parameter",
        "Created By",
        "Created By Parameter",
        "Last Modified By",
        "Modified By Parameter",
        "Created On",
        "Modified On",
        "Thumbnail Path",
        "Warnings",
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)

    category_colors: dict[str, str] = {}
    next_color = 0

    for row in rows:
        md = row.metadata
        mass_each = format_mass(row.mass_kg_each, report_unit)
        mass_total = format_mass(row.total_mass_kg, report_unit)

        excel_row = [
            "\n".join(row.indices),
            row.model,
            report_unit,
            mass_each,
            row.qty,
            mass_total,
            row.model_type,
            row.immediate_assembly,
            row.immediate_assembly_index,
            row.assembly_path,
            row.depth,
            row.top_qty,
            md.native_mass_unit,
            md.material,
            row.category,
            md.owner,
            md.owner_parameter,
            md.created_by,
            md.created_by_parameter,
            md.modified_by,
            md.modified_by_parameter,
            md.created_on,
            md.modified_on,
            md.thumbnail_path,
            warnings_text(row.warnings),
        ]
        ws.append(excel_row)
        r = ws.max_row

        # Base alignment and borders.
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="E7E6E6"))

        if row.model_type == "ASSEMBLY":
            core = row.depth == 1 and not unique_sheet
            fill = depth_fill(row.depth, core=core)
            font_color = "FFFFFF" if core else "1F1F1F"
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).fill = fill
                ws.cell(r, c).font = Font(
                    color=font_color,
                    bold=(c in (1, 2, 6, 7)),
                )
        else:
            # Category-specific text emphasis.
            if row.category != "OTHER":
                if row.category not in category_colors:
                    category_colors[row.category] = CATEGORY_FONT_COLORS[
                        next_color % len(CATEGORY_FONT_COLORS)
                    ]
                    next_color += 1
                color = category_colors[row.category]
                for c in (1, 2, 15):
                    ws.cell(r, c).font = Font(color=color, bold=True)

        if row.warnings:
            ws.cell(r, 25).fill = PatternFill("solid", fgColor="FFF2CC")
            ws.cell(r, 25).font = Font(color="9C6500", bold=True)

        if md.thumbnail_path:
            path_cell = ws.cell(r, 24)
            path_cell.hyperlink = Path(md.thumbnail_path).as_uri()
            path_cell.style = "Hyperlink"
            if settings.thumbnails:
                try:
                    img = XLImage(md.thumbnail_path)
                    img.width = 96
                    img.height = 72
                    # Anchor over the Thumbnail Path cell; keep the path text too.
                    ws.add_image(img, f"X{r}")
                    ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 15, 58)
                except Exception:
                    # Path remains available even if Pillow/openpyxl image embedding fails.
                    pass

        # Highlight invalid mass cells.
        if mass_each is None:
            ws.cell(r, 4).value = "Inf"
            ws.cell(r, 4).fill = PatternFill("solid", fgColor="F4CCCC")
        if mass_total is None:
            ws.cell(r, 6).value = "Inf"
            ws.cell(r, 6).fill = PatternFill("solid", fgColor="F4CCCC")

    # Number formats.
    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row_idx, 4).number_format = "0.000000"
        ws.cell(row_idx, 6).number_format = "0.000000"
        ws.cell(row_idx, 5).number_format = "0"
        ws.cell(row_idx, 11).number_format = "0"
        ws.cell(row_idx, 12).number_format = "0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Create an Excel table if non-empty.
    if ws.max_row >= 2:
        table_name = re.sub(r"\W+", "", sheet_name) + "Table"
        tab = Table(displayName=table_name[:255], ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False,
        )
        ws.add_table(tab)

    widths = {
        1: 22,
        2: 32,
        3: 12,
        4: 14,
        5: 8,
        6: 14,
        7: 12,
        8: 30,
        9: 24,
        10: 46,
        11: 8,
        12: 16,
        13: 18,
        14: 28,
        15: 18,
        16: 24,
        17: 20,
        18: 26,
        19: 22,
        20: 26,
        21: 22,
        22: 22,
        23: 22,
        24: 46,
        25: 42,
    }
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws.row_dimensions[1].height = 34


def write_warnings_sheet(wb: Workbook, warnings: Sequence[WarningRecord]) -> None:
    ws = wb.create_sheet("Warnings")
    ws.append(["Severity", "Code", "Model", "Index", "Message"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="C00000")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    for wrn in sorted(
        warnings,
        key=lambda w: (
            SEVERITY_ORDER.get(w.severity, 99),
            w.code,
            w.model.casefold(),
            natural_path_key(w.index or "0"),
        ),
    ):
        ws.append([wrn.severity, wrn.code, wrn.model, wrn.index, wrn.message])
        r = ws.max_row
        fill = {
            "ERROR": "F4CCCC",
            "WARNING": "FFF2CC",
            "INFO": "D9EAF7",
        }.get(wrn.severity, "FFFFFF")
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=fill)
        for c in range(1, 6):
            ws.cell(r, c).alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, width in enumerate([12, 30, 34, 20, 80], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def leaf_category_masses(root: Occurrence) -> dict[str, float]:
    totals: dict[str, float] = collections.defaultdict(float)
    invalid: set[str] = set()
    for occ in iter_occurrences(root):
        if not occ.is_part:
            continue
        if math.isfinite(occ.mass_kg):
            totals[occ.category] += occ.mass_kg
        else:
            invalid.add(occ.category)
    for category in invalid:
        totals[category] = math.inf
    return dict(totals)


def top_level_mass_breakdown(root: Occurrence) -> list[tuple[str, float]]:
    """
    Core-subassembly/direct-child breakdown with FASTENER mass removed from each
    top-level bucket, then added once as a separate global bucket.
    """
    result: list[tuple[str, float]] = []
    grouped: dict[str, float] = collections.defaultdict(float)
    invalid: set[str] = set()

    for child in root.children:
        if child.is_part and child.category == "FASTENER":
            continue
        value = mass_excluding_category_kg(child, "FASTENER")
        if math.isfinite(value):
            grouped[child.model] += value
        else:
            invalid.add(child.model)

    for model in sorted(grouped, key=str.casefold):
        result.append((model, grouped[model]))
    for model in sorted(invalid, key=str.casefold):
        result.append((model, math.inf))

    fastener_mass = 0.0
    fastener_invalid = False
    for occ in iter_occurrences(root):
        if occ.is_part and occ.category == "FASTENER":
            if math.isfinite(occ.mass_kg):
                fastener_mass += occ.mass_kg
            else:
                fastener_invalid = True
    if fastener_mass or fastener_invalid:
        result.append(("FASTENER (all leaf parts)", math.inf if fastener_invalid else fastener_mass))
    return result


def write_summary_sheet(
    wb: Workbook,
    root: Occurrence,
    report_unit: str,
    categories: collections.OrderedDict[str, list[str]],
) -> None:
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = APP_NAME
    ws["A2"] = f"Version {APP_VERSION}"
    ws["A4"] = "Top Assembly"
    ws["B4"] = root.model
    ws["A5"] = "Report Mass Unit"
    ws["B5"] = report_unit
    ws["A6"] = "Derived Assembly Mass"
    ws["B6"] = format_mass(root.mass_kg, report_unit) if math.isfinite(root.mass_kg) else "Inf"
    ws["B6"].number_format = "0.000000"

    title_fill = PatternFill("solid", fgColor="1F4E78")
    for c in ("A1", "A2"):
        ws[c].font = Font(color="FFFFFF", bold=True, size=14 if c == "A1" else 10)
        ws[c].fill = title_fill
    ws.merge_cells("A1:H1")
    ws.merge_cells("A2:H2")

    top_breakdown = top_level_mass_breakdown(root)
    start = 10
    ws.cell(start, 1, "Top-Level Mass Breakdown")
    ws.cell(start, 2, f"Mass [{report_unit}]")
    ws.cell(start, 1).font = ws.cell(start, 2).font = Font(bold=True, color="FFFFFF")
    ws.cell(start, 1).fill = ws.cell(start, 2).fill = PatternFill("solid", fgColor="5B9BD5")

    valid_top_rows = []
    for label, value_kg in top_breakdown:
        start += 1
        ws.cell(start, 1, label)
        if math.isfinite(value_kg):
            ws.cell(start, 2, kg_to_mass(value_kg, report_unit))
            ws.cell(start, 2).number_format = "0.000000"
            valid_top_rows.append(start)
        else:
            ws.cell(start, 2, "Inf")

    if valid_top_rows:
        chart = PieChart()
        data = Reference(ws, min_col=2, min_row=min(valid_top_rows), max_row=max(valid_top_rows))
        labels = Reference(ws, min_col=1, min_row=min(valid_top_rows), max_row=max(valid_top_rows))
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.title = "Top-Level Mass Distribution\n(Fasteners Separated)"
        chart.height = 9
        chart.width = 13
        ws.add_chart(chart, "D4")

    cat_start = max(start + 3, 25)
    ws.cell(cat_start, 1, "Leaf Part Category")
    ws.cell(cat_start, 2, f"Mass [{report_unit}]")
    ws.cell(cat_start, 1).font = ws.cell(cat_start, 2).font = Font(bold=True, color="FFFFFF")
    ws.cell(cat_start, 1).fill = ws.cell(cat_start, 2).fill = PatternFill("solid", fgColor="70AD47")

    cat_masses = leaf_category_masses(root)
    valid_cat_rows = []
    r = cat_start
    for category in sorted(cat_masses, key=str.casefold):
        r += 1
        ws.cell(r, 1, category)
        value_kg = cat_masses[category]
        if math.isfinite(value_kg):
            ws.cell(r, 2, kg_to_mass(value_kg, report_unit))
            ws.cell(r, 2).number_format = "0.000000"
            valid_cat_rows.append(r)
        else:
            ws.cell(r, 2, "Inf")

    if valid_cat_rows:
        bar = BarChart()
        data = Reference(ws, min_col=2, min_row=min(valid_cat_rows), max_row=max(valid_cat_rows))
        labels = Reference(ws, min_col=1, min_row=min(valid_cat_rows), max_row=max(valid_cat_rows))
        bar.add_data(data, titles_from_data=False)
        bar.set_categories(labels)
        bar.title = "Leaf-Part Mass by Category"
        bar.y_axis.title = f"Mass [{report_unit}]"
        bar.x_axis.title = "Category"
        bar.height = 8
        bar.width = 13
        ws.add_chart(bar, f"D{cat_start}")

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 20
    ws.freeze_panes = "A4"


def write_metadata_sheet(
    wb: Workbook,
    settings: Settings,
    root: Occurrence,
    report_unit: str,
) -> None:
    ws = wb.create_sheet("Report Metadata")
    rows = [
        ("Application", APP_NAME),
        ("Version", APP_VERSION),
        ("Top Assembly", root.model),
        ("Report Unit", report_unit),
        ("Mass Warning Threshold", settings.high_mass_warning),
        ("CREOSON Host", settings.host),
        ("CREOSON Port", settings.port),
        ("Creo Version Hint", settings.creo_version),
        ("Thumbnails Enabled", settings.thumbnails),
        ("Thumbnail Directory", str(settings.thumbnail_dir.resolve())),
        ("Owner Parameters", ", ".join(settings.owner_parameters)),
        ("Creator Parameters", ", ".join(settings.creator_parameters)),
        ("Modifier Parameters", ", ".join(settings.modifier_parameters)),
        ("Created-On Parameters", ", ".join(settings.created_on_parameters)),
        ("Modified-On Parameters", ", ".join(settings.modified_on_parameters)),
        (
            "Categories",
            json.dumps(settings.categories, indent=2),
        ),
        (
            "Hierarchy De-duplication",
            (
                "Identical model names are combined only when they share the same "
                "immediate parent occurrence path. Global de-duplication is shown "
                "separately on the Unique Rollup sheet."
            ),
        ),
        (
            "Assembly Mass Method",
            (
                "Assembly masses are recursively derived from descendant occurrence "
                "masses. Part masses come from Creo mass properties and are converted "
                "to a common report unit before summation."
            ),
        ),
    ]
    ws.append(["Setting", "Value"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.font = Font(color="FFFFFF", bold=True)
    for key, value in rows:
        ws.append([key, value])
        ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 100


def save_workbook(
    output: Path,
    root: Occurrence,
    hierarchy_rows: Sequence[GroupRow],
    unique_rows: Sequence[GroupRow],
    report_unit: str,
    settings: Settings,
    warnings: Sequence[WarningRecord],
) -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    write_summary_sheet(wb, root, report_unit, settings.categories)
    write_bom_sheet(
        wb,
        "BOM Hierarchy",
        hierarchy_rows,
        report_unit,
        settings,
        unique_sheet=False,
    )
    write_bom_sheet(
        wb,
        "Unique Rollup",
        unique_rows,
        report_unit,
        settings,
        unique_sheet=True,
    )
    write_warnings_sheet(wb, warnings)
    write_metadata_sheet(wb, settings, root, report_unit)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def load_config(path: Optional[Path]) -> dict[str, Any]:
    if not path:
        return {}
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("Configuration JSON must contain an object at the top level.")
    return data


def ordered_categories(value: Any) -> collections.OrderedDict[str, list[str]]:
    if value is None:
        return collections.OrderedDict([("FASTENER", ["NAS"]), ("HARNESS", ["HARNESS"])])
    if isinstance(value, dict):
        result = collections.OrderedDict()
        for name, patterns in value.items():
            if isinstance(patterns, str):
                patterns = [patterns]
            result[str(name).upper()] = [str(x) for x in patterns]
        return result
    raise ValueError("'categories' must be an object mapping category names to pattern lists.")


def merge_settings(args: argparse.Namespace, config: dict[str, Any]) -> Settings:
    settings = Settings()

    # Config first.
    for field_name in dataclasses.asdict(settings):
        if field_name in config and field_name != "categories":
            value = config[field_name]
            if field_name in ("output", "thumbnail_dir"):
                value = Path(value)
            setattr(settings, field_name, value)
    if "categories" in config:
        settings.categories = ordered_categories(config["categories"])

    # CLI overrides.
    for name in (
        "host",
        "port",
        "assembly",
        "creo_version",
        "report_unit",
        "high_mass_warning",
    ):
        value = getattr(args, name, None)
        if value is not None:
            setattr(settings, name, value)

    if args.output is not None:
        settings.output = Path(args.output)
    if args.thumbnail_dir is not None:
        settings.thumbnail_dir = Path(args.thumbnail_dir)
    if args.thumbnails:
        settings.thumbnails = True
    if args.include_skeletons:
        settings.include_skeletons = True

    if args.owner_parameter:
        settings.owner_parameters = args.owner_parameter
    if args.creator_parameter:
        settings.creator_parameters = args.creator_parameter
    if args.modifier_parameter:
        settings.modifier_parameters = args.modifier_parameter
    if args.thumbnail_view:
        settings.thumbnail_views = args.thumbnail_view

    if args.category:
        # CLI category entries replace config/default categories if provided.
        parsed = collections.OrderedDict()
        for item in args.category:
            if "=" not in item:
                raise ValueError(
                    f"Invalid --category '{item}'. Use NAME=PATTERN1,PATTERN2"
                )
            name, raw_patterns = item.split("=", 1)
            patterns = [p.strip() for p in raw_patterns.split(",") if p.strip()]
            if not name.strip() or not patterns:
                raise ValueError(
                    f"Invalid --category '{item}'. Use NAME=PATTERN1,PATTERN2"
                )
            parsed[name.strip().upper()] = patterns
        settings.categories = parsed

    # Ensure FASTENER exists if user defined categories but omitted it.
    if "FASTENER" not in settings.categories:
        console_warn(
            "No FASTENER category is configured. Fastener-separated summary "
            "charts will therefore have no dedicated fastener bucket."
        )

    settings.output = settings.output.resolve()
    settings.thumbnail_dir = settings.thumbnail_dir.resolve()
    return settings


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Create a hierarchy-aware Creo assembly mass/material/ownership "
            "report through creopyson / CREOSON."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--assembly", help="Top-level .asm file. Defaults to active Creo model.")
    p.add_argument("--output", help="Output XLSX file.")
    p.add_argument("--config", type=Path, help="Optional JSON configuration file.")
    p.add_argument("--host", help="CREOSON host.")
    p.add_argument("--port", type=int, help="CREOSON port.")
    p.add_argument("--creo-version", type=int, help="Creo major version hint for CREOSON.")
    p.add_argument(
        "--report-unit",
        help="Common output mass unit (kg, g, mg, lb/lbm, oz, tonne).",
    )
    p.add_argument(
        "--high-mass-warning",
        type=float,
        help="Warn when an individual PART mass exceeds this value in report units.",
    )
    p.add_argument("--thumbnails", action="store_true", help="Export and embed JPEG thumbnails.")
    p.add_argument("--thumbnail-dir", help="Directory for exported model JPEGs.")
    p.add_argument(
        "--thumbnail-view",
        action="append",
        help=(
            "Preferred saved Creo view name. Repeat to specify priority. "
            "Defaults: ISO, ISOMETRIC, TRIMETRIC, DEFAULT."
        ),
    )
    p.add_argument("--include-skeletons", action="store_true", help="Include skeleton models.")
    p.add_argument(
        "--owner-parameter",
        action="append",
        help="Owner parameter in priority order. Repeat option as needed.",
    )
    p.add_argument(
        "--creator-parameter",
        action="append",
        help="Creator parameter in priority order. Repeat option as needed.",
    )
    p.add_argument(
        "--modifier-parameter",
        action="append",
        help="Last-modifier parameter in priority order. Repeat option as needed.",
    )
    p.add_argument(
        "--category",
        action="append",
        help=(
            "Filename classification: NAME=PATTERN1,PATTERN2. Repeat for multiple "
            "categories. First matching category wins. Example: "
            "--category FASTENER=NAS,MS,AN --category HARNESS=HARNESS,WH"
        ),
    )
    p.add_argument(
        "--debug",
        action="store_true",
        help="Print a full traceback if the report fails.",
    )
    return p


def print_report_summary(
    root: Occurrence,
    hierarchy_rows: Sequence[GroupRow],
    unique_rows: Sequence[GroupRow],
    warnings: Sequence[WarningRecord],
    output: Path,
    report_unit: str,
) -> None:
    counts = collections.Counter(w.severity for w in warnings)
    total_occ = sum(1 for _ in iter_occurrences(root))
    mass_value = (
        f"{kg_to_mass(root.mass_kg, report_unit):.6g} {report_unit}"
        if math.isfinite(root.mass_kg)
        else f"Inf {report_unit}"
    )

    console_info("-" * 72)
    console_info(f"Top assembly: {root.model}")
    console_info(f"Component occurrences: {total_occ}")
    console_info(f"Hierarchy rows after parent-local de-duplication: {len(hierarchy_rows)}")
    console_info(f"Unique models: {len(unique_rows)}")
    console_info(f"Derived assembly mass: {mass_value}")
    console_info(
        "Warnings: "
        f"{counts.get('ERROR', 0)} error(s), "
        f"{counts.get('WARNING', 0)} warning(s), "
        f"{counts.get('INFO', 0)} info item(s)"
    )
    console_info(f"Excel report: {output}")
    console_info("-" * 72)


def run(settings: Settings) -> Path:
    warnings: list[WarningRecord] = []
    adapter = CreoAdapter(settings)

    console_info(f"{APP_NAME} v{APP_VERSION}")
    console_info(f"Connecting to CREOSON at {settings.host}:{settings.port} ...")
    adapter.connect()
    console_info("Connected.")

    if settings.assembly:
        top_model = settings.assembly
        console_info(f"Opening requested top assembly: {top_model}")
        adapter.open_model(top_model)
    else:
        active = adapter.get_active()
        top_model = normalize_model_name(active.get("file"))
        if not top_model:
            raise RuntimeError(
                "No --assembly was supplied and Creo has no active model."
            )

    if not top_model.lower().endswith(".asm"):
        raise ValueError(
            f"Top model '{top_model}' is not an .asm file. "
            "This report expects a top-level assembly."
        )

    console_info(f"Reading BOM hierarchy for {top_model} ...")
    bom = adapter.get_bom(top_model)
    root = parse_bom_tree(bom, top_model)
    occurrences = list(iter_occurrences(root))
    if not occurrences:
        console_warn("The assembly BOM contained no component occurrences.")

    unique_models = sorted({o.model for o in occurrences}, key=str.casefold)
    console_info(
        f"Parsed {len(occurrences)} occurrence(s) across "
        f"{len(unique_models)} unique model(s)."
    )

    metadata = load_model_metadata(adapter, unique_models, settings, warnings)
    report_unit = infer_report_unit(
        adapter, top_model, metadata, settings.report_unit, warnings
    )
    console_info(f"Using common report mass unit: {report_unit}")

    validate_part_metadata(metadata, settings, report_unit, warnings)
    assign_occurrence_properties(root, metadata, settings.categories)
    add_occurrence_warning_indices(root, metadata, warnings)

    if settings.thumbnails:
        generate_thumbnails(
            adapter,
            unique_models,
            metadata,
            settings,
            warnings,
            restore_model=top_model,
        )

    console_info("Building hierarchy-preserving and global rollups ...")
    hierarchy_rows = group_hierarchy_rows(root, metadata)
    unique_rows = group_unique_rows(root, metadata)

    # Attach row paths to model-level warnings where possible for easier triage.
    first_path_by_model: dict[str, str] = {}
    for occ in occurrences:
        first_path_by_model.setdefault(occ.model, occ.path)
    for wrn in warnings:
        if not wrn.index:
            wrn.index = first_path_by_model.get(wrn.model, "")

    console_info("Writing formatted Excel workbook and charts ...")
    save_workbook(
        settings.output,
        root,
        hierarchy_rows,
        unique_rows,
        report_unit,
        settings,
        warnings,
    )
    print_report_summary(
        root,
        hierarchy_rows,
        unique_rows,
        warnings,
        settings.output,
        report_unit,
    )
    return settings.output


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
        settings = merge_settings(args, config)
        run(settings)
        return 0
    except KeyboardInterrupt:
        console_warn("Cancelled by user.")
        return 130
    except Exception as exc:
        console_error(str(exc))
        if getattr(args, "debug", False):
            traceback.print_exc()
        else:
            console_error("Re-run with --debug for a full traceback.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
