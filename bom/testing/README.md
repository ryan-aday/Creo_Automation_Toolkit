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
  - Owner.
  - Creator / last modifier.
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

---

# Version 3: Wiring Harness Mass Analysis

Version 3 adds `harness_mass_report.py`, a separate Cabling-focused workflow.
The base assembly report also writes a machine-readable JSON sidecar beside the
XLSX workbook, for example:

```text
MAIN_ASSEMBLY_mass_report.xlsx
MAIN_ASSEMBLY_mass_report.json
```

## Cleaner person fields

The BOM worksheets now show the values people care about:

- `Owner`
- `Created By`
- `Last Modified By`
- `Created On`
- `Modified On`

The repetitive `Owner Parameter`, `Created By Parameter`, and
`Modified By Parameter` columns have been removed from the BOM worksheets. The
script can still use a prioritized list of Creo/Windchill parameter names
internally to locate those values.

## Harness discovery

By default the harness script searches assembly names for `_ROUT`. Additional or
replacement tokens can be supplied in `harness_config.example.json` or from the
command line:

```powershell
python .\harness_mass_report.py .\MAIN_ASSEMBLY_mass_report.json `
  --token _ROUT `
  --token _HARNESS
```

The source can be either the JSON sidecar or the Excel workbook. Only matching
`.asm` models are selected.

## Generating one XML per harness

Creo Parametric Cabling can export logical data as Creo Schematics XML. CREOSON
does not expose a dedicated Cabling XML-export endpoint, so the automation uses
`interface_mapkey()`.

Record a mapkey in your own Creo release that performs:

```text
Cabling -> Logical Data -> Export -> Creo Schematic
```

Save the mapkey script body in a text file and replace the recorded output path
with this exact placeholder:

```text
{xml_path}
```

Then set `mapkey_template` in `harness_config.example.json`, or copy
`cabling_export_mapkey.example.txt` and use it as your starting point.

Run:

```powershell
python .\harness_mass_report.py .\MAIN_ASSEMBLY_mass_report.json `
  --config .\harness_config.json `
  --export-xml `
  --excel .\MAIN_ASSEMBLY_mass_report.xlsx
```

The script displays/activates each discovered routing assembly, substitutes a
unique XML destination into the mapkey, executes the mapkey through CREOSON,
and waits for the XML to appear in the configured XML directory.

If the XML files already exist, omit `--export-xml`; the remaining analysis is
fully automatic.

## XML parsing model

The parser uses Python's standard `xml.etree.ElementTree` library and converts
XML element attributes/parameter entries into normalized dictionaries. It is
intentionally schema-tolerant because Creo/Schematics XML varies between
versions and site conventions.

It searches spool records for fields such as:

- `SPOOL_NAME`
- `TYPE` / `OBJ_TYPE`
- `WIRE_GAUGE`
- `THICKNESS`, `DIAMETER`, or `OUTER_DIAMETER`
- `DENSITY`
- `MASS_UNITS`
- `UNITS`
- insulation/material/construction metadata

It searches connection records for:

- connection/wire/cable name
- spool reference
- routed length
- length units
- from/to references when present

## Wire and cable mass equations

### Creo spool linear-density mode — default

PTC defines Cabling spool `DENSITY` as a **linear density** in mass per unit
length. Accordingly, the default calculation is:

```text
m_wire = lambda * L
```

where:

- `m_wire` = wire/cable mass
- `lambda` = spool linear density
- `L` = routed connection length

Diameter is **not multiplied into this equation**, because it is already
implicitly represented by the spool's linear density.

### Volumetric-density mode

For custom XML where density represents material density instead of Creo linear
density, set:

```json
"density_mode": "volumetric"
```

The calculation becomes:

```text
A = pi*d^2/4
m_wire = rho*A*L
```

where `rho` is volumetric density and `d` is diameter.

### Gauge fallback

If XML does not contain `THICKNESS` / diameter but contains an AWG gauge, the
script derives the **bare conductor diameter** from:

```text
d_mm = 0.127 * 92^((36-AWG)/39)
```

This is deliberately reported as an `AWG bare-conductor fallback`, because the
insulated outside diameter can be materially larger. Explicit Creo spool
`THICKNESS` is preferred whenever present.

## Tape mass

When tape is only cosmetic in Creo, a rule can specify width, thickness,
material density, overlap, and either a length override or a factor applied to
the routed connection length.

For helical tape wrapping:

```text
pitch = width * (1 - overlap)
L_tape = L_axial * sqrt(1 + (pi*D/pitch)^2)
V_tape = L_tape * width * thickness
m_tape = rho_tape * V_tape
```

An areal-density input may be used instead of volumetric density.

## Overbraid mass

For an assumed thin overbraid shell:

```text
A_surface = pi*D*L
V_braid ~= A_surface * thickness * coverage
m_braid = rho_braid * V_braid
```

Alternatively, if an areal density is known:

```text
m_braid = areal_density * A_surface * coverage
```

The `coverage_fraction` is useful because braided sleeving is not a fully dense
solid cylindrical shell.

## Cosmetic-rule example

See `harness_config.example.json`. A rule can be scoped independently by both
harness and connection wildcard patterns:

```json
{
  "name": "HARNESS_TAPE",
  "kind": "tape",
  "harness_pattern": "*_ROUT*",
  "connection_pattern": "PWR_*",
  "thickness_mm": 0.13,
  "width_mm": 19.0,
  "material_density_kg_m3": 1200.0,
  "overlap_fraction": 0.5,
  "length_factor": 1.0
}
```

A fixed cosmetic application length can be supplied with `length_m` when the
covering length does not match the routed connection length.

## Harness workbook sheets

Running the harness analyzer against an Excel report adds/replaces:

### `Harness Masses`

A detailed connection-level table containing:

- source assembly index
- harness assembly
- immediate owning assembly
- owner
- creator
- last modifier
- connection name
- spool/covering name
- item type
- wire gauge
- effective diameter
- density mode and normalized density
- routed/application length
- calculated mass
- from/to references
- material/construction field
- assumptions / diameter source
- XML source file
- warnings

Tape and overbraid assumption rows are visually separated from ordinary
wire/cable rows.

### `Harness Summary`

One row per routing assembly containing:

- harness/model name and source indices
- immediate assembly
- owner/creator/modifier
- XML path
- number of parsed spools
- number of connections
- calculated total harness mass
- unresolved row count
- warnings

A harness-mass bar chart is added to the summary sheet.

## Suggested workflow

```powershell
# 1. Build the ordinary assembly report + JSON sidecar.
python .\creo_assembly_mass_report.py `
  --assembly MAIN_ASSEMBLY.asm `
  --output .\MAIN_ASSEMBLY_mass_report.xlsx

# 2. Discover routing assemblies, export their XML, calculate mass,
#    and attach Harness Masses / Harness Summary to the workbook.
python .\harness_mass_report.py .\MAIN_ASSEMBLY_mass_report.json `
  --config .\harness_config.json `
  --export-xml `
  --excel .\MAIN_ASSEMBLY_mass_report.xlsx
```

For the first validation run, compare at least one known harness against Creo
Cabling mass properties and manually check the XML spool density units and
connection lengths. Site-specific spool libraries sometimes use nonstandard
unit strings; add aliases if your installation does.
