#!/usr/bin/env python3
"""Export the active Creo assembly's nested part masses to CSV/XLSX.

Run this on Windows while Creo Parametric is open with a top-level assembly active.
The script uses PTC's Creo VB API through pywin32; it never saves a model or edits geometry.
"""

from __future__ import annotations

import argparse
import csv
import gc
import hashlib
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Iterator


G_STANDARD_M_S2 = 9.80665
NEWTONS_PER_LBF = 4.4482216152605

OCCURRENCE_FIELDS = [
    "occurrence_path",
    "component_id_path",
    "level",
    "part_file",
    "model_full_name",
    "source_path",
    "common_name",
    "generic_name",
    "instance_name",
    "quantity",
    "mass",
    "mass_unit",
    "unit_system",
    "mass_kg",
    "weight_N",
    "weight_lbf",
    "volume",
    "density",
    "density_warning",
    "mass_basis",
    "thumbnail_file",
]

ROLLUP_FIELDS = [
    "part_file",
    "model_full_name",
    "source_path",
    "common_name",
    "generic_name",
    "instance_name",
    "quantity",
    "mass_each",
    "mass_unit",
    "total_mass_native",
    "unit_system",
    "mass_each_kg",
    "total_mass_kg",
    "total_weight_N",
    "total_weight_lbf",
    "volume_each",
    "density",
    "density_warning",
    "mass_basis",
    "thumbnail_file",
]

WARNING_FIELDS = ["occurrence_path", "code", "message"]


@dataclass(frozen=True)
class MassData:
    mass: float | None
    volume: float | None
    density: float | None
    mass_unit: str
    unit_system: str
    mass_kg: float | None
    density_warning: str


def read_member(obj: Any, name: str, default: Any = "") -> Any:
    """Read a COM property (or a zero-argument COM method) without crashing."""
    try:
        value = getattr(obj, name)
        return value() if callable(value) else value
    except Exception:
        return default


def float_or_none(value: Any) -> float | None:
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def com_items(sequence: Any) -> Iterator[Any]:
    """Iterate PTC sequence objects, which are normally zero-indexed COM lists."""
    if sequence is None:
        return
    try:
        count = int(sequence.Count)
        for index in range(count):
            yield sequence.Item(index)
        return
    except Exception:
        pass
    try:
        yield from sequence
    except TypeError:
        return


def descriptor_extension(descriptor: Any) -> str:
    try:
        extension = descriptor.GetExtension()
    except Exception:
        extension = read_member(descriptor, "Extension", "")
    return str(extension or "").lower().lstrip(".")


def descriptor_file_name(descriptor: Any) -> str:
    try:
        return str(descriptor.GetFileName())
    except Exception:
        return str(read_member(descriptor, "FileName", ""))


def model_full_name(model: Any, fallback: str) -> str:
    for member in ("FullName", "FileName", "InstanceName"):
        value = str(read_member(model, member, "") or "").strip()
        if value:
            return value
    return fallback


def model_key(model: Any, fallback: str) -> str:
    origin = str(read_member(model, "Origin", "") or "").strip()
    return f"{origin}|{model_full_name(model, fallback)}".casefold()


def get_model(session: Any, descriptor: Any) -> Any:
    model = None
    try:
        model = session.GetModelFromDescr(descriptor)
    except Exception:
        pass
    if model is None:
        model = session.RetrieveModel(descriptor)
    return model


def list_component_features(solid: Any) -> list[Any]:
    """Return active component features without relying on generated enum wrappers."""
    try:
        features = solid.ListFeaturesByType(False, None)
    except Exception as exc:
        raise RuntimeError(f"Could not list assembly features: {exc}") from exc

    components: list[Any] = []
    for feature in com_items(features):
        type_name = str(read_member(feature, "FeatTypeName", "") or "").casefold()
        if type_name == "component":
            components.append(feature)
            continue
        # Some Creo releases localize/omit FeatTypeName. A component feature is
        # also identifiable by the ModelDescr property.
        try:
            descriptor = feature.ModelDescr
            if descriptor is not None and descriptor_extension(descriptor) in {"asm", "prt"}:
                components.append(feature)
        except Exception:
            pass
    return components


def normalized_unit_name(unit: Any) -> str:
    candidates = (
        read_member(unit, "Expression", ""),
        read_member(unit, "Name", ""),
        read_member(unit, "FullName", ""),
    )
    for candidate in candidates:
        text = str(candidate or "").strip()
        if text:
            return text
    return "unknown"


def direct_mass_unit_to_kg(name: str) -> float | None:
    key = re.sub(r"[\s_\-()]", "", name.casefold())
    known = {
        "kg": 1.0,
        "kilogram": 1.0,
        "kilograms": 1.0,
        "g": 1e-3,
        "gram": 1e-3,
        "grams": 1e-3,
        "mg": 1e-6,
        "milligram": 1e-6,
        "lb": 0.45359237,
        "lbm": 0.45359237,
        "pound": 0.45359237,
        "pounds": 0.45359237,
        "oz": 0.028349523125,
        "ozm": 0.028349523125,
        "ounce": 0.028349523125,
        "slug": 14.593902937206,
        "t": 1000.0,
        "tonne": 1000.0,
        "metricton": 1000.0,
    }
    if key in known:
        return known[key]
    # Creo may expose expressions such as "kg^1".
    if key.endswith("^1"):
        return known.get(key[:-2])
    return None


def unit_to_kg_factor(unit: Any, seen: set[str] | None = None) -> float | None:
    """Resolve a Creo mass unit to kilograms through its reference-unit chain."""
    if unit is None:
        return None
    name = normalized_unit_name(unit)
    direct = direct_mass_unit_to_kg(name)
    if direct is not None:
        return direct

    seen = set() if seen is None else seen
    identity = f"{name}|{read_member(unit, 'FullName', '')}".casefold()
    if identity in seen:
        return None
    seen.add(identity)

    reference = read_member(unit, "ReferenceUnit", None)
    conversion = read_member(unit, "ConversionFactor", None)
    scale = float_or_none(read_member(conversion, "Scale", None))
    if reference is None or scale in (None, 0.0):
        return None
    reference_to_kg = unit_to_kg_factor(reference, seen)
    if reference_to_kg is None:
        return None

    # PTC defines: actual-unit value = Scale * reference-unit value + Offset.
    # Mass-unit conversions should have zero offset, so actual -> reference is /Scale.
    return reference_to_kg / scale


def get_mass_data(part: Any) -> MassData:
    try:
        # EpfcMP_DENSITY_DEFAULT is enum value 0. This method respects the
        # material density and, unlike GetMassProperty, does not update the
        # model's stored mass-property state.
        properties = part.GetMassPropertyWithDensity(None, 0, 0.0)
    except Exception:
        try:
            properties = part.GetMassProperty(None)
        except Exception:
            # Some generated COM wrappers reject None for a nullable BSTR.
            properties = part.GetMassProperty("")

    mass = float_or_none(read_member(properties, "Mass", None))
    volume = float_or_none(read_member(properties, "Volume", None))
    density = float_or_none(read_member(properties, "Density", None))

    unit_system = part.GetPrincipalUnits()
    unit_system_name = str(read_member(unit_system, "Name", "") or "unknown")
    # EpfcUNIT_MASS is value 1 (EpfcUNIT_LENGTH is value 0).
    mass_unit_object = unit_system.GetUnit(1)
    mass_unit_name = normalized_unit_name(mass_unit_object)
    kg_factor = unit_to_kg_factor(mass_unit_object)
    mass_kg = mass * kg_factor if mass is not None and kg_factor is not None else None

    density_warning = ""
    if density is not None and math.isclose(density, 1.0, rel_tol=0.0, abs_tol=1e-12):
        density_warning = "Check material assignment: Creo returned the default density 1.0"

    return MassData(
        mass=mass,
        volume=volume,
        density=density,
        mass_unit=mass_unit_name,
        unit_system=unit_system_name,
        mass_kg=mass_kg,
        density_warning=density_warning,
    )


def dispatch_first(program_ids: Iterable[str]) -> Any:
    from win32com.client.dynamic import Dispatch as DynamicDispatch

    errors: list[str] = []
    for program_id in program_ids:
        try:
            return DynamicDispatch(program_id)
        except Exception as exc:
            errors.append(f"{program_id}: {exc}")
    raise RuntimeError("No compatible Creo COM class was available (" + "; ".join(errors) + ")")


def connect_to_creo(timeout_seconds: int) -> Any:
    factory = dispatch_first(("pfcls.pfcAsyncConnection", "pfcls.CCpfcAsyncConnection"))
    errors: list[str] = []
    # The nullable TextPath parameter is exposed differently by different pywin32/
    # Creo combinations, so use the documented null first and two safe fallbacks.
    for text_path in (None, "", "."):
        try:
            return factory.Connect("", "", text_path, timeout_seconds)
        except Exception as exc:
            errors.append(f"TextPath={text_path!r}: {exc}")
    raise RuntimeError("Could not connect to a running Creo session. " + " | ".join(errors))


def safe_filename(value: str, limit: int = 70) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._") or "part"
    return stem[:limit]


def export_thumbnail(
    session: Any,
    model: Any,
    destination: Path,
    view_name: str,
    width_inches: float,
    height_inches: float,
) -> None:
    original_window = read_member(session, "CurrentWindow", None)
    window = None
    created_window = False
    try:
        try:
            window = session.GetModelWindow(model)
        except Exception:
            window = None
        if window is None:
            window = session.CreateModelWindow(model)
            created_window = True

        window.Activate()
        model.Display()

        # A user-supplied/saved view gives repeatable thumbnails. If it is absent,
        # use Creo's DEFAULT view and finally whatever view is already active.
        if view_name:
            try:
                model.RetrieveView(view_name)
            except Exception:
                try:
                    model.RetrieveView("DEFAULT")
                except Exception:
                    pass
        window.Repaint()

        factory = dispatch_first(
            (
                "pfcls.pfcJPEGImageExportInstructions",
                "pfcls.CCpfcJPEGImageExportInstructions",
            )
        )
        instructions = factory.Create(width_inches, height_inches)
        window.ExportRasterImage(str(destination), instructions)
        if not destination.exists():
            raise RuntimeError("Creo returned without creating the JPEG")
    finally:
        # A newly created window cannot be closed while it is current.
        if original_window is not None:
            try:
                original_window.Activate()
            except Exception:
                pass
        if created_window and window is not None:
            try:
                window.Close()
            except Exception:
                pass


def part_metadata(model: Any, fallback_name: str) -> dict[str, str]:
    return {
        "part_file": str(read_member(model, "FileName", "") or fallback_name),
        "model_full_name": model_full_name(model, fallback_name),
        "source_path": str(read_member(model, "Origin", "") or ""),
        "common_name": str(read_member(model, "CommonName", "") or ""),
        "generic_name": str(read_member(model, "GenericName", "") or ""),
        "instance_name": str(read_member(model, "InstanceName", "") or ""),
    }


def collect_occurrences(
    session: Any,
    root_assembly: Any,
    root_name: str,
    output_dir: Path,
    make_thumbnails: bool,
    view_name: str,
    thumbnail_width: float,
    thumbnail_height: float,
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    rows: list[dict[str, Any]] = []
    warnings: list[dict[str, str]] = []
    mass_cache: dict[str, MassData] = {}
    thumbnail_cache: dict[str, str] = {}
    thumbnail_dir = output_dir / "thumbnails"
    if make_thumbnails:
        thumbnail_dir.mkdir(parents=True, exist_ok=True)

    def warning(path: str, code: str, message: str) -> None:
        warnings.append({"occurrence_path": path, "code": code, "message": message})

    def visit_assembly(
        assembly: Any,
        names: tuple[str, ...],
        ids: tuple[int, ...],
        ancestor_assemblies: tuple[str, ...],
    ) -> None:
        parent_path = " > ".join(names)
        try:
            components = list_component_features(assembly)
        except Exception as exc:
            warning(parent_path, "FEATURE_LIST_FAILED", str(exc))
            return

        for feature in components:
            feature_id_value = read_member(feature, "Id", -1)
            try:
                feature_id = int(feature_id_value)
            except (TypeError, ValueError):
                feature_id = -1

            try:
                descriptor = feature.ModelDescr
                extension = descriptor_extension(descriptor)
                child_file = descriptor_file_name(descriptor) or f"component_{feature_id}"
            except Exception as exc:
                occurrence = f"{parent_path} > component[{feature_id}]"
                warning(
                    occurrence,
                    "COMPONENT_UNAVAILABLE",
                    "Component descriptor is unavailable (often suppressed, excluded, or missing): "
                    + str(exc),
                )
                continue

            child_label = f"{child_file}[{feature_id}]"
            occurrence_names = names + (child_label,)
            occurrence_ids = ids + (feature_id,)
            occurrence_path = " > ".join(occurrence_names)

            try:
                child_model = get_model(session, descriptor)
                if child_model is None:
                    raise RuntimeError("Creo returned no model object")
            except Exception as exc:
                warning(occurrence_path, "MODEL_RETRIEVE_FAILED", str(exc))
                continue

            identity = model_key(child_model, child_file)
            if extension == "asm":
                if identity in ancestor_assemblies:
                    warning(occurrence_path, "ASSEMBLY_CYCLE", "Recursive assembly reference was skipped")
                    continue
                visit_assembly(
                    child_model,
                    occurrence_names,
                    occurrence_ids,
                    ancestor_assemblies + (identity,),
                )
                continue
            if extension != "prt":
                warning(occurrence_path, "UNSUPPORTED_COMPONENT", f"Skipped .{extension or '?'} component")
                continue

            if identity not in mass_cache:
                try:
                    mass_cache[identity] = get_mass_data(child_model)
                except Exception as exc:
                    warning(occurrence_path, "MASS_PROPERTY_FAILED", str(exc))
                    mass_cache[identity] = MassData(None, None, None, "", "", None, "")
            mass = mass_cache[identity]

            thumbnail_relative = ""
            if make_thumbnails:
                if identity not in thumbnail_cache:
                    suffix = hashlib.sha1(identity.encode("utf-8")).hexdigest()[:10]
                    thumbnail_name = f"{safe_filename(Path(child_file).stem)}_{suffix}.jpg"
                    destination = thumbnail_dir / thumbnail_name
                    try:
                        export_thumbnail(
                            session,
                            child_model,
                            destination,
                            view_name,
                            thumbnail_width,
                            thumbnail_height,
                        )
                        thumbnail_cache[identity] = destination.relative_to(output_dir).as_posix()
                    except Exception as exc:
                        warning(occurrence_path, "THUMBNAIL_FAILED", str(exc))
                        thumbnail_cache[identity] = ""
                thumbnail_relative = thumbnail_cache[identity]

            metadata = part_metadata(child_model, child_file)
            row: dict[str, Any] = {
                "occurrence_path": occurrence_path,
                "component_id_path": "/".join(str(value) for value in occurrence_ids),
                "level": len(occurrence_ids),
                **metadata,
                "quantity": 1,
                "mass": mass.mass,
                "mass_unit": mass.mass_unit,
                "unit_system": mass.unit_system,
                "mass_kg": mass.mass_kg,
                "weight_N": mass.mass_kg * G_STANDARD_M_S2 if mass.mass_kg is not None else None,
                "weight_lbf": (
                    mass.mass_kg * G_STANDARD_M_S2 / NEWTONS_PER_LBF
                    if mass.mass_kg is not None
                    else None
                ),
                "volume": mass.volume,
                "density": mass.density,
                "density_warning": mass.density_warning,
                "mass_basis": "part_model",
                "thumbnail_file": thumbnail_relative,
            }
            rows.append(row)

    root_identity = model_key(root_assembly, root_name)
    visit_assembly(root_assembly, (root_name,), (), (root_identity,))
    return rows, warnings


def roll_up(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        key = (
            str(row["source_path"]).casefold(),
            str(row["model_full_name"]).casefold(),
            str(row["mass_unit"]).casefold(),
        )
        groups[key].append(row)

    result: list[dict[str, Any]] = []
    for grouped_rows in groups.values():
        first = grouped_rows[0]
        quantity = len(grouped_rows)
        mass_each = float_or_none(first["mass"])
        mass_each_kg = float_or_none(first["mass_kg"])
        result.append(
            {
                "part_file": first["part_file"],
                "model_full_name": first["model_full_name"],
                "source_path": first["source_path"],
                "common_name": first["common_name"],
                "generic_name": first["generic_name"],
                "instance_name": first["instance_name"],
                "quantity": quantity,
                "mass_each": mass_each,
                "mass_unit": first["mass_unit"],
                "total_mass_native": mass_each * quantity if mass_each is not None else None,
                "unit_system": first["unit_system"],
                "mass_each_kg": mass_each_kg,
                "total_mass_kg": mass_each_kg * quantity if mass_each_kg is not None else None,
                "total_weight_N": (
                    mass_each_kg * quantity * G_STANDARD_M_S2 if mass_each_kg is not None else None
                ),
                "total_weight_lbf": (
                    mass_each_kg * quantity * G_STANDARD_M_S2 / NEWTONS_PER_LBF
                    if mass_each_kg is not None
                    else None
                ),
                "volume_each": first["volume"],
                "density": first["density"],
                "density_warning": first["density_warning"],
                "mass_basis": first["mass_basis"],
                "thumbnail_file": first["thumbnail_file"],
            }
        )
    result.sort(key=lambda row: str(row["model_full_name"]).casefold())
    return result


def write_csv_file(path: Path, fields: list[str], rows: list[dict[str, Any]]) -> None:
    # UTF-8 with BOM opens cleanly in localized versions of Excel.
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def add_worksheet(
    workbook: Any,
    title: str,
    fields: list[str],
    rows: list[dict[str, Any]],
    output_dir: Path,
    embed_images: bool,
) -> None:
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    sheet = workbook.create_sheet(title)
    spreadsheet_fields = (["thumbnail"] + fields) if "thumbnail_file" in fields else fields
    sheet.append(spreadsheet_fields)

    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.freeze_panes = "A2"

    for row_number, row in enumerate(rows, start=2):
        values = ([""] if "thumbnail_file" in fields else []) + [row.get(field, "") for field in fields]
        sheet.append(values)
        if embed_images and row.get("thumbnail_file"):
            image_path = output_dir / str(row["thumbnail_file"])
            if image_path.is_file():
                image = XLImage(str(image_path))
                image.width = 96
                image.height = 72
                image.anchor = f"A{row_number}"
                sheet.add_image(image)
                sheet.row_dimensions[row_number].height = 58

    sheet.auto_filter.ref = sheet.dimensions
    for column_index, field in enumerate(spreadsheet_fields, start=1):
        if field == "thumbnail":
            width = 15
        elif field in {"occurrence_path", "density_warning", "message"}:
            width = 55
        elif field in {"model_full_name", "source_path", "thumbnail_file"}:
            width = 34
        else:
            width = min(max(len(field) + 2, 13), 22)
        sheet.column_dimensions[get_column_letter(column_index)].width = width


def write_workbook(
    path: Path,
    occurrence_rows: list[dict[str, Any]],
    rollup_rows: list[dict[str, Any]],
    warnings: list[dict[str, str]],
    output_dir: Path,
    embed_images: bool,
) -> None:
    from openpyxl import Workbook

    workbook = Workbook()
    workbook.remove(workbook.active)
    add_worksheet(workbook, "Occurrences", OCCURRENCE_FIELDS, occurrence_rows, output_dir, embed_images)
    add_worksheet(workbook, "Rollup", ROLLUP_FIELDS, rollup_rows, output_dir, embed_images)
    add_worksheet(workbook, "Warnings", WARNING_FIELDS, warnings, output_dir, False)
    workbook.save(path)


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Export every active nested part occurrence in Creo's current assembly."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path.cwd() / "creo_mass_report",
        help="Output directory (default: ./creo_mass_report)",
    )
    parser.add_argument(
        "--thumbnails",
        action="store_true",
        help="Export part JPEGs and embed them in the XLSX workbook",
    )
    parser.add_argument(
        "--view",
        default="ISOVIEW",
        help="Saved Creo view for thumbnails; falls back to DEFAULT (default: ISOVIEW)",
    )
    parser.add_argument("--thumbnail-width", type=float, default=1.5, help="JPEG width in inches")
    parser.add_argument("--thumbnail-height", type=float, default=1.125, help="JPEG height in inches")
    parser.add_argument("--timeout", type=int, default=5, help="Creo connection timeout in seconds")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    if sys.platform != "win32":
        print("ERROR: Creo VB API automation requires Windows.", file=sys.stderr)
        return 2
    if args.thumbnail_width <= 0 or args.thumbnail_height <= 0:
        print("ERROR: Thumbnail dimensions must be positive.", file=sys.stderr)
        return 2

    try:
        import pythoncom
        import win32com.client  # noqa: F401 - verifies that pywin32 is installed
    except ImportError:
        print("ERROR: pywin32 is not installed. Run: py -m pip install -r requirements.txt", file=sys.stderr)
        return 2

    output_dir = args.output.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    connection = None
    pythoncom.CoInitialize()
    try:
        connection = connect_to_creo(args.timeout)
        session = connection.Session
        root_model = session.CurrentModel
        if root_model is None:
            raise RuntimeError("Creo has no active model. Open the top-level assembly first.")

        root_name = model_full_name(root_model, "active_model")
        if Path(root_name).suffix.casefold() != ".asm":
            raise RuntimeError(f"The active Creo model is not an assembly: {root_name}")

        occurrence_rows, warnings = collect_occurrences(
            session=session,
            root_assembly=root_model,
            root_name=root_name,
            output_dir=output_dir,
            make_thumbnails=args.thumbnails,
            view_name=args.view,
            thumbnail_width=args.thumbnail_width,
            thumbnail_height=args.thumbnail_height,
        )
        rollup_rows = roll_up(occurrence_rows)

        occurrences_csv = output_dir / "parts_occurrences.csv"
        rollup_csv = output_dir / "parts_rollup.csv"
        warnings_csv = output_dir / "warnings.csv"
        workbook_path = output_dir / "parts_with_thumbnails.xlsx"
        write_csv_file(occurrences_csv, OCCURRENCE_FIELDS, occurrence_rows)
        write_csv_file(rollup_csv, ROLLUP_FIELDS, rollup_rows)
        write_csv_file(warnings_csv, WARNING_FIELDS, warnings)
        write_workbook(
            workbook_path,
            occurrence_rows,
            rollup_rows,
            warnings,
            output_dir,
            args.thumbnails,
        )

        print(f"Assembly: {root_name}")
        print(f"Part occurrences: {len(occurrence_rows)}")
        print(f"Unique parts: {len(rollup_rows)}")
        print(f"Warnings: {len(warnings)}")
        print(f"CSV: {occurrences_csv}")
        print(f"Rollup CSV: {rollup_csv}")
        print(f"Workbook: {workbook_path}")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if connection is not None:
            try:
                # Disconnect only this automation client; Creo keeps running.
                connection.Disconnect(args.timeout)
            except Exception:
                pass
        connection = None
        gc.collect()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
