# Creo nested-part mass report

This Windows script connects to an already-running Creo Parametric session through
PTC's supported VB API, walks every active nested subassembly, and reports every
part occurrence. It does not save models or edit geometry/parameters.

It creates:

- `parts_occurrences.csv` — one row per part occurrence, including the complete
  assembly path and component-ID path.
- `parts_rollup.csv` — one row per unique part with quantity and total mass.
- `warnings.csv` — suppressed/excluded/missing components and any mass/image errors.
- `parts_with_thumbnails.xlsx` — the same occurrence and rollup data in an Excel
  workbook. With `--thumbnails`, JPEGs are embedded in its cells.
- `thumbnails/*.jpg` — one image per unique part when `--thumbnails` is enabled.

CSV is plain text and cannot contain embedded images. Its `thumbnail_file` column
contains a relative path; the images themselves are embedded in the XLSX file.

## 1. Enable Creo's VB API once

The VB API must be installed with Creo Parametric.

1. Set the `PRO_COMM_MSG_EXE` environment variable to Creo's
   `Common Files\\<machine type>\\obj\\pro_comm_msg.exe` (typically
   `x86_win64`; use the exact path from your Creo installation).
2. Run `<Creo load point>\\Parametric\\bin\\vb_api_register.bat`. An elevated
   Command Prompt may be required.
3. Keep Python and Creo at the same architecture (normally both 64-bit).

PTC setup reference:
https://support.ptc.com/help/creo_toolkit/vbapi_plus/usascii/creo_toolkit/user_guide/VBAPIExamples.html

## 2. Install Python dependencies

From this folder in PowerShell or Command Prompt:

```powershell
py -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -r requirements.txt
```

## 3. Generate the report

1. Start Creo Parametric.
2. Open and regenerate the top-level assembly.
3. Activate the simplified representation you want reported.
4. Confirm materials/densities are assigned.
5. Run:

```powershell
.venv\Scripts\python creo_assembly_mass_report.py --output C:\Temp\creo_mass_report
```

Add rendered part thumbnails:

```powershell
.venv\Scripts\python creo_assembly_mass_report.py `
  --output C:\Temp\creo_mass_report `
  --thumbnails `
  --view ISOVIEW
```

`ISOVIEW` should be a saved view in each part. When it is absent, the script tries
`DEFAULT` and then leaves the current view unchanged. Thumbnail export briefly
activates/creates model windows, restores the original window, and does not save.

Run `py creo_assembly_mass_report.py --help` for image-size and timeout options.

## What the numbers mean

- `mass` is the value returned by Creo for the part model in its principal mass unit.
- `mass_kg` is converted through Creo's unit/reference-unit chain when possible.
- `weight_N` and `weight_lbf` use standard gravity, 9.80665 m/s².
- `mass_basis` is `part_model`: repeated occurrences use the standalone part model's
  mass. This is correct for ordinary rigid occurrences. Occurrence-specific assembly
  cuts or flexible-component geometry require `IpfcAssembly.GetMassPropertyByCompPath`
  instead.
- Creo returns density `1.0` when no material density is assigned. Such rows are
  marked in `density_warning`; treat those masses as suspect until verified.

Only components active in the current simplified representation can be traversed.
Suppressed, excluded, or missing components are recorded in `warnings.csv` when Creo
exposes enough information to identify them.

## Troubleshooting

- **`Invalid class string` / COM class unavailable:** rerun
  `vb_api_register.bat`, verify `PRO_COMM_MSG_EXE`, and restart Creo and the shell.
- **Cannot connect:** keep one Creo session open with the assembly active. Check that
  `pro_comm_msg.exe` matches the running Creo release.
- **Every density is 1.0:** assign a material/density and regenerate mass properties.
- **No `mass_kg`:** the original Creo mass and unit are still in the report, but the
  custom unit chain could not be resolved automatically.
- **Blank thumbnails:** test a saved view name, keep graphics enabled, and inspect
  `warnings.csv`. The CSV remains valid even when a JPEG export fails.

## PTC API references

- Assembly components and nested structure:
  https://support.ptc.com/help/creo_toolkit/vbapi_plus/usascii/creo_toolkit/user_guide/Structure_of_Assemblies_and_Assembly_Objects.html
- Solid mass properties:
  https://support.ptc.com/help/creo_toolkit/vbapi_plus/usascii/creo_toolkit/user_guide/Mass_Properties.html
- JPEG raster export:
  https://support.ptc.com/help/creo_toolkit/vbapi_plus/usascii/creo_toolkit/api/dita/t-pfcWindow-Window.html
