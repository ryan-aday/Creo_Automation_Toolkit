# Creo Assembly Mass Report

A console-first Python reporting tool for **PTC Creo Parametric** using
**creopyson + CREOSON**. It converts the live assembly BOM into a structured,
hierarchy-aware Excel workbook with mass, material, ownership, creator/modifier
metadata, warning flags, optional thumbnails, and summary charts.

## What this version changes

The central design rule is that de-duplication is **parent-local**, not global.

If four identical bolts exist in the full assembly, but two occur under
`SUBASM_A` and two under `SUBASM_C`, the `BOM Hierarchy` sheet keeps one bolt
row under `SUBASM_A` with `Qty = 2` and a second bolt row under `SUBASM_C` with
`Qty = 2`. Their component indices are combined in the `Index` cell.

A separate `Unique Rollup` sheet performs the global de-duplication by model
name and lists all occurrence paths and parent assemblies.

## Main workbook sheets

- **Summary**
  - Derived top-level assembly mass.
  - Pie chart of top-level mass, with leaf fasteners removed from the
    subassembly buckets and shown as a separate fastener bucket.
  - Bar chart of leaf-part mass by filename category.
- **BOM Hierarchy**
  - Index / combined occurrence indices.
  - Model name.
  - Mass units.
  - Individual mass.
  - Quantity under the same immediate parent occurrence.
  - Total mass.
  - Type (assembly or part).
  - Immediate owning assembly and its occurrence index.
  - Full assembly path.
  - Quantity in the full top-level assembly.
  - Native mass unit.
  - Material.
  - Filename category.
  - Owner and the parameter that supplied it.
  - Creator / last modifier and source parameter.
  - Windchill created/modified timestamps where exposed.
  - Thumbnail path.
  - Warning codes.
- **Unique Rollup**
  - One row per unique model name across the entire BOM.
  - Combined occurrence indices and all parent assemblies.
- **Warnings**
  - Mass failures.
  - Unsupported/missing mass units.
  - Non-positive or unusually high mass.
  - Default / missing material.
  - Missing owner metadata.
  - Thumbnail failures.
  - Derived assembly mass failures caused by descendant data.
- **Report Metadata**
  - Configuration and report methodology.

## Assembly mass method

Part mass comes from `file_massprops()`.

Assembly mass is **not** taken from `file_massprops()` because some Creo /
CREOSON combinations can return unusable values for assemblies. Instead, each
assembly occurrence is recursively evaluated as:

```text
assembly occurrence mass = sum(direct child occurrence masses)
```

Since child assembly masses are themselves derived recursively, this produces a
top-level assembly mass from the leaf parts while preserving quantities.

All recognized part masses are first converted to kilograms internally and then
converted to the chosen report unit. This prevents mixed model units from being
summed directly.

Supported built-in units:

- kg
- g
- mg
- lb / lbm
- oz
- tonne

Extend `MASS_UNIT_TO_KG` in the script if your site returns a different unit
string.

## Ownership and Windchill metadata

Default owner priority:

```text
OWNER
DESIGNER
DRAWN_BY
PTC_WM_CREATED_BY
```

Default creator priority:

```text
PTC_WM_CREATED_BY
CREATED_BY
CREATOR
```

Default modifier priority:

```text
PTC_WM_MODIFIED_BY
MODIFIED_BY
LAST_MODIFIED_BY
```

For Windchill-managed models, `PTC_WM_CREATED_BY`,
`PTC_WM_MODIFIED_BY`, `PTC_WM_CREATED_ON`, and `PTC_WM_MODIFIED_ON` are
available only when those system attributes are mapped/exposed to Creo. CREOSON
itself provides limited Windchill workspace operations; this script therefore
uses the model parameters that Creo can see instead of attempting a direct
Windchill database query.

## Thumbnail behavior

With `--thumbnails`, each unique model is displayed and exported once as a JPEG.

The script searches saved model views in this default priority:

1. `ISO`
2. `ISOMETRIC`
3. `TRIMETRIC`
4. `DEFAULT`
5. First available saved view

The image is embedded in Excel and the original JPEG path is retained as a
clickable link.

Because this is based on the saved model view, use your site's normal ISO view
name via repeated `--thumbnail-view` options if necessary.

## Installation

Use the Python installation that can communicate with your local CREOSON server.

```bash
python -m pip install -r requirements.txt
```

You must already have:

1. Creo Parametric running or startable.
2. CREOSON configured for that Creo installation.
3. The assembly and required dependencies available to the Creo session /
   workspace.

## Basic usage

Use the current active Creo assembly:

```bash
python creo_assembly_mass_report.py
```

Specify an assembly and output file:

```bash
python creo_assembly_mass_report.py \
  --assembly MAIN_ASSEMBLY.asm \
  --output reports/main_assembly_mass.xlsx
```

On Windows PowerShell:

```powershell
python .\creo_assembly_mass_report.py `
  --assembly MAIN_ASSEMBLY.asm `
  --output .\reports\main_assembly_mass.xlsx
```

With thumbnails and Creo 11:

```powershell
python .\creo_assembly_mass_report.py `
  --assembly MAIN_ASSEMBLY.asm `
  --creo-version 11 `
  --thumbnails `
  --thumbnail-view ISO `
  --thumbnail-view ISOMETRIC
```

## Custom filename categories

The first matching category wins.

For example:

```powershell
python .\creo_assembly_mass_report.py `
  --category "FASTENER=NAS,MS,AN" `
  --category "ASE=ASE" `
  --category "HARNESS=HARNESS,WIRE"
```

The `FASTENER` category has special handling in the summary pie chart: leaf
fastener mass is removed from core-subassembly buckets and added once as its own
bucket.

If you define categories from the command line, those entries replace the
default category set.

## Custom owner / creator / modifier parameters

```powershell
python .\creo_assembly_mass_report.py `
  --owner-parameter RESPONSIBLE_ENGINEER `
  --owner-parameter OWNER `
  --creator-parameter PTC_WM_CREATED_BY `
  --modifier-parameter PTC_WM_MODIFIED_BY
```

Repeated options define priority order.

## JSON configuration

Copy `config.example.json` and customize it:

```powershell
python .\creo_assembly_mass_report.py `
  --config .\config.example.json `
  --assembly MAIN_ASSEMBLY.asm
```

Command-line arguments override matching JSON settings.

## Notes on warnings

- `Inf` in the Excel mass columns is intentional: it means the source mass could
  not be read or safely converted.
- A non-finite leaf-part mass propagates upward to every containing assembly.
- `MATERIAL_DEFAULT_OR_MISSING` is raised if the current material is blank,
  contains `DEFAULT`, or contains `_NO_MATL_ASSIGNED`.
- The default high-mass threshold is `50` in the chosen report unit.

## CREOSON / creopyson API calls used

The implementation uses public creopyson client methods including:

- `bom_get_paths(..., paths=True)`
- `file_massprops()`
- `file_get_mass_units()`
- `file_get_cur_material()`
- `parameter_list()`
- `view_list()`
- `view_activate()`
- `interface_export_image()`

## Recommended validation on your installation

Before using the report for release/signoff, validate one known assembly with:

1. A repeated fastener under one subassembly.
2. The same fastener under a second subassembly.
3. A nested subassembly.
4. A deliberately missing/default material.
5. A model with known `PTC_WM_CREATED_BY` / `PTC_WM_MODIFIED_BY` values.
6. A known part mass in each mass unit your organization uses.
7. A known top-level mass from Creo for comparison with the recursive rollup.

CREOSON behavior can vary with Creo version, Windchill mapping, simplified
representations, and local model standards, so this validation step is
important.
