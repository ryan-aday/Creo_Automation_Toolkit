# Creo Geometry Quality Audit

A Python 3.10+ design-screening tool that opens a user-selected Creo Parametric part or assembly through `creopyson`/CREOSON, regenerates it, exports a STEP snapshot, and performs geometry and metadata QA with Trimesh.

It includes both a Streamlit review interface and a command-line runner. The 3D viewer renders ordinary bodies translucent blue-gray; bodies involved in confirmed interference are opaque red; evaluated clearance locations are drawn as line segments.

## Checks included

| Check | Method | Result |
|---|---|---|
| Body interference | AABB broad phase, then Manifold3D positive-volume Boolean; optional `python-fcl` fallback | Body pair, overlap volume when available, method/confidence note |
| Global minimum gap | AABB pruning followed by bidirectional vertex/triangle-center-to-surface proximity | Body pair, approximate closest points, triangle and reconstructed patch IDs |
| Pair-specific gap/tolerance | Case-insensitive regular-expression body matching; optional reconstructed patch IDs; min/max or target ± tolerance | Nominal gap and pass/fail for each matching occurrence pair |
| Worst-case scalar stack | `stack_closure` is subtracted from measured clearance | Effective clearance used for rule pass/fail |
| Thermal clearance | Uniform isotropic scale about each body centroid or the global origin | Hot/cold-state gap-rule results when both bodies have verified CTE values |
| Missing/default material | Creo current material name plus optional `_NO_MATL_SET` model parameter | One finding per source part |
| Failed/missing-reference symptoms | Creo regeneration response and `UNREGENERATED`/`INACTIVE` feature states | Feature/model details returned by CREOSON |
| Small surface patches | Connected coplanar triangle facets; ungrouped curved triangles remain individual patches | Patch area, patch type, triangle count, centroid |
| Mesh health | Watertightness | Accuracy warning for affected bodies |

## Important scope limits

This application is deliberately explicit about what it can and cannot prove.

1. **The geometry checks use tessellation.** A STEP B-rep is converted to triangles by Trimesh/Cascadio. Results close to the export chord-height/angular tolerances must be confirmed in Creo.
2. **Mesh “faces” are not guaranteed to be Creo faces.** Planar connected triangles are grouped into patches. A curved analytic face will generally be represented by many residual triangles. Use the returned closest point and body names to locate the critical region in Creo.
3. **Negative fallback collision results can be inconclusive.** Positive-volume Manifold booleans are preferred. If a mesh is not watertight or the Boolean fails, the tool uses `python-fcl` when installed and finally sampled containment. The report records the method.
4. **CREOSON does not expose every Creo material-file property uniformly.** The tool reads the assigned material name and looks for configured model parameter aliases such as `CTE`. Otherwise, it uses the user-maintained material table. It never silently invents CTE.
5. **Thermal expansion is a screening model, not constrained thermoelastic FEA.** It assumes a spatially uniform temperature, small-strain linear isotropic expansion, and a specified free-expansion anchor.
6. **Missing references are detected indirectly.** CREOSON exposes regeneration errors and `UNREGENERATED` feature state. It does not provide a complete dependency-graph/reference diagnostic equivalent to Creo's Reference Viewer. A clean result is not proof that all external references are intentional and stable.

For released hardware, use this report to focus the native Creo Global Interference, Measure, Reference Viewer, model-check, tolerance-analysis, and FEA workflows.

## Prerequisites

- Windows workstation with Creo Parametric running.
- CREOSON server running with its Creo connection established.
- Python 3.10 or newer. Trimesh 5.1 currently supports Python 3.10+.
- The Python process and CREOSON/Creo must see the model and export paths under the same filesystem names.

## Installation

```powershell
cd creo_geometry_audit
py -3.12 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Optional faster collision/distance backend:

```powershell
python -m pip install -r requirements-optional.txt
```

`python-fcl` wheel availability varies by Python version and platform. The primary positive-volume interference backend is Manifold3D, so the application remains usable without FCL.

## Start the Streamlit interface

```powershell
streamlit run app.py
```

Or double-click `run_app.bat` after activating/configuring the intended Python environment.

Enter the complete path to a `.prt` or `.asm` visible to Creo, review the assumptions before inputs, configure checks and body-pair rules, and select **Run Creo audit**.

## Command-line use

Minimal run:

```powershell
creo-audit "C:\Creo_Work\gearbox.asm" --creo-version 11 --minimum-gap 0.25 --minimum-face-area 0.05
```

Configuration-driven run:

```powershell
creo-audit --config config.example.json
```

The process returns exit code `2` when the audit contains one or more error-severity findings, otherwise `0`. Connection, input, or runtime failures produce Python's normal nonzero failure status.

## Gap-rule definitions

Each rule matches all occurrence-level body pairs whose STEP/body name or mapped Creo source-model name matches `body_a` and `body_b` as regular expressions.

- `minimum`, `maximum`: allowed clearance window.
- `patch_a`, `patch_b`: optional reconstructed surface-patch IDs from a prior scan. Leave null to search the complete bodies. Patch IDs are tied to that exact tessellation/export and may change after regeneration or export-setting changes.
- `target`, `tolerance`: when both are supplied, they replace the min/max window with `target - tolerance` through `target + tolerance`.
- `stack_closure`: scalar worst-case closure allowance subtracted from nominal or thermal mesh clearance.
- `description`: free-text requirement/rationale included in the report.

The evaluated effective clearance is

\[
g_{\mathrm{eff}}=\max\left(0,\;g_{\mathrm{mesh}}-C_{\mathrm{stack}}\right).
\]

This scalar closure model is useful when a controlled worst-case stack result already exists. It does not calculate a full dimensional loop from arbitrary Creo dimensions. A full stack requires datum/loop definitions, contributor signs, distributions, and correlations.

## Thermal method

For body vertex position \(\mathbf{x}\), expansion anchor \(\mathbf{o}\), coefficient of thermal expansion \(\alpha\), and temperature change \(\Delta T\), the screening transform is

\[
\mathbf{x}_{T}=\mathbf{o}+\left(1+\alpha\Delta T\right)\left(\mathbf{x}-\mathbf{o}\right).
\]

The transformed bodies are then passed through the same gap-rule scan. `centroid` is a neutral free-expansion visualization assumption. `global_origin` is useful only when the assembly origin is the physically meaningful shared anchor. Constrained parts usually need explicit constraint-aware displacement fields or thermoelastic FEA.

The default `cte_parameter_scale` is `1e-6`, meaning a Creo model parameter value of `23.6` is interpreted as \(23.6\times10^{-6}/\mathrm{K}\). Set the scale to `1.0` if the model parameter is already stored in `1/K`. Values in the material table are always entered directly in `1/K`.

## Outputs

Every run writes:

- `<model>_geometry_audit.step`: analysis snapshot exported by Creo.
- `<model>_audit.html`: standalone interactive 3D report.
- `<model>_audit.json`: complete machine-readable configuration, findings, closest points, methods, and warnings.
- `<model>_findings.csv`: flattened findings table.

## Project layout

```text
app.py                         Streamlit application
creo_audit/creo_bridge.py     CREOSON connection, regeneration, metadata, STEP export
creo_audit/mesh_io.py         STEP scene loading, unit conversion, occurrence/body extraction
creo_audit/checks.py          collision, clearance, and surface-patch algorithms
creo_audit/thermal.py         isotropic thermal transform
creo_audit/engine.py          audit orchestration
creo_audit/visualization.py   transparent/red Plotly 3D scene
creo_audit/reporting.py       HTML, JSON, and CSV exports
creo_audit/cli.py             command-line entry point
tests/                        synthetic geometry verification
```

## Technical references

- [Creopyson 0.7.8 documentation](https://creopyson.readthedocs.io/en/latest/) — client connection and Creo automation API.
- [Creopyson package API](https://creopyson.readthedocs.io/en/latest/creopyson.html) — `file_open`, `file_regenerate`, `feature_list`, `bom_get_paths`, material and parameter calls.
- [Creopyson interface source](https://creopyson.readthedocs.io/en/latest/_modules/creopyson/interface.html) — supported STEP export API and options.
- [Trimesh 5.1 documentation](https://trimesh.org/) — scene graphs, transforms, proximity, Boolean operations, and visualization data.
- [Trimesh Cascadio STEP loader](https://trimesh.org/trimesh.exchange.cascade.html) — OpenCASCADE-backed STEP-to-scene tessellation.
- [PTC STEP export overview](https://support.ptc.com/help/creo/creo_pma/r12/usascii/data_exchange/interface/About_Exporting_Parts_and_Assemblies_to_STEP.html) — Creo STEP export profiles and geometry options.
