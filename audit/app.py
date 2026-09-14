from __future__ import annotations

import io
import math
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from creo_audit.engine import AuditEngine
from creo_audit.models import AuditConfig, GapRule, MaterialSpec
from creo_audit.reporting import write_reports
from creo_audit.visualization import build_figure


st.set_page_config(page_title="Creo Geometry Audit", page_icon="🧭", layout="wide")
st.title("Creo Geometry Quality Audit")
st.caption("CREOSON extraction • STEP tessellation • interference, clearance, material, feature, and small-patch review")

with st.expander("Read before running — scope and assumptions", expanded=True):
    st.warning(
        "This is a design-screening tool. Geometry is analyzed from a tessellated STEP snapshot, so "
        "clearance and patch areas are resolution-dependent. Confirm critical results with Creo's native "
        "B-rep Measure/Global Interference tools and a controlled export profile."
    )
    st.markdown(
        """
- Creo Parametric and CREOSON must already be running and connected on the computer that runs this app.
- The entered `.prt`/`.asm` path is a path visible to Creo—not a browser upload.
- Thermal expansion is linear, uniform, isotropic, and free about the selected anchor. It does not model constraints, temperature gradients, creep, or nonlinear CTE.
- A gap-rule `stack_closure` is subtracted from the measured gap as a scalar worst-case allowance.
- “Faces” are reconstructed planar facets plus individual residual triangles. Curved CAD-face identity is not preserved by STEP tessellation.
- CREOSON reliably exposes assigned material names. CTE is used only from an explicit model parameter or the editable table below.
"""
    )

with st.sidebar:
    st.header("Creo connection")
    model_path = st.text_input("Creo model path", placeholder=r"C:\projects\assembly.asm")
    host = st.text_input("CREOSON host", "localhost")
    port = st.number_input("CREOSON port", min_value=1, max_value=65535, value=9056, step=1)
    creo_version_text = st.text_input("Creo major version (optional)", placeholder="11")
    output_dir = st.text_input("Audit output directory", "audit_output")

    st.header("Checks")
    run_interference = st.checkbox("Body interference", True)
    run_global_gap = st.checkbox("Global minimum gap", True)
    minimum_gap = st.number_input("Minimum global gap", min_value=0.0, value=0.0, format="%.8g")
    run_materials = st.checkbox("Material assignment", True)
    run_features = st.checkbox("Failed/unregenerated features", True)
    run_small_faces = st.checkbox("Small tessellated patches", True)
    minimum_face_area = st.number_input("Minimum patch area", min_value=0.0, value=0.0, format="%.8g")
    distance_samples = st.number_input("Distance sample budget per body", 100, 100_000, 2500, 100)

    with st.expander("Metadata conventions"):
        default_material_text = st.text_input(
            "Unset/default material names",
            ", DEFAULT, _NO_MATL_SET, NO_MATERIAL",
            help="Comma-separated; a blank first entry treats null/empty material as unset.",
        )
        material_tag_parameter = st.text_input("No-material Boolean/tag parameter", "_NO_MATL_SET")
        cte_parameter_text = st.text_input(
            "CTE model parameter aliases",
            "CTE, ALPHA, THERMAL_EXPANSION_COEFFICIENT, COEFF_THERMAL_EXPANSION",
        )
        cte_parameter_scale = st.number_input(
            "CTE model parameter scale",
            min_value=0.0,
            value=1.0e-6,
            format="%.8g",
            help="Use 1e-6 when Creo stores CTE as microstrain/K; use 1 when it stores 1/K.",
        )

    st.header("Thermal state")
    thermal_enabled = st.checkbox("Evaluate hot/cold gap rules", False)
    reference_temperature = st.number_input("Reference temperature", value=20.0)
    operating_temperature = st.number_input("Operating temperature", value=80.0, disabled=not thermal_enabled)
    thermal_anchor = st.selectbox("Expansion anchor", ["centroid", "global_origin"], disabled=not thermal_enabled)

st.subheader("Pair-specific clearance rules")
st.caption("Body columns are case-insensitive regular expressions matched against STEP body and Creo source-model names.")
gap_frame = st.data_editor(
    pd.DataFrame(
        [
            {
                "body_a": "",
                "body_b": "",
                "patch_a": None,
                "patch_b": None,
                "minimum": 0.0,
                "maximum": None,
                "target": None,
                "tolerance": None,
                "stack_closure": 0.0,
                "description": "",
            }
        ]
    ),
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="gap_rules",
)

st.subheader("Material CTE table")
st.caption("Values are coefficients in 1/K. Replace representative defaults with approved, temperature-appropriate data.")
material_frame = st.data_editor(
    pd.DataFrame(
        [
            {"material": "AL_6061_T6", "cte_per_k": 23.6e-6, "source": "representative; verify specification"},
            {"material": "TI_6AL_4V", "cte_per_k": 8.6e-6, "source": "representative; verify specification"},
            {"material": "STEEL", "cte_per_k": 12.0e-6, "source": "representative; grade-dependent"},
        ]
    ),
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="materials",
)


def optional_float(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    return float(value)


def build_config() -> AuditConfig:
    rules = []
    for row in gap_frame.to_dict("records"):
        if not str(row.get("body_a", "")).strip() or not str(row.get("body_b", "")).strip():
            continue
        rules.append(
            GapRule(
                body_a=str(row["body_a"]), body_b=str(row["body_b"]),
                patch_a=None if optional_float(row.get("patch_a")) is None else int(row["patch_a"]),
                patch_b=None if optional_float(row.get("patch_b")) is None else int(row["patch_b"]),
                minimum=float(row.get("minimum") or 0.0), maximum=optional_float(row.get("maximum")),
                target=optional_float(row.get("target")), tolerance=optional_float(row.get("tolerance")),
                stack_closure=float(row.get("stack_closure") or 0.0), description=str(row.get("description") or ""),
            )
        )
    materials = {}
    for row in material_frame.to_dict("records"):
        name = str(row.get("material") or "").strip()
        cte = optional_float(row.get("cte_per_k"))
        if name:
            materials[name] = MaterialSpec(name=name, cte_per_k=cte, source=str(row.get("source") or ""))
    version = int(creo_version_text) if creo_version_text.strip() else None
    return AuditConfig(
        model_path=model_path, creoson_host=host, creoson_port=int(port), creo_version=version,
        export_directory=output_dir, minimum_global_gap=float(minimum_gap), minimum_face_area=float(minimum_face_area),
        distance_samples=int(distance_samples), reference_temperature=float(reference_temperature),
        operating_temperature=float(operating_temperature) if thermal_enabled else None,
        thermal_anchor=thermal_anchor, material_library=materials, gap_rules=rules,
        default_material_names=[item.strip() for item in default_material_text.split(",")],
        material_tag_parameter=material_tag_parameter.strip(),
        cte_parameter_names=[item.strip() for item in cte_parameter_text.split(",") if item.strip()],
        cte_parameter_scale=float(cte_parameter_scale),
        run_interference=run_interference, run_global_gap=run_global_gap, run_small_faces=run_small_faces,
        run_materials=run_materials, run_feature_health=run_features,
    )


if st.button("Run Creo audit", type="primary", use_container_width=True):
    if not model_path.strip():
        st.error("Enter the Creo .prt or .asm path.")
    else:
        try:
            config = build_config()
            with st.status("Running Creo audit…", expanded=True) as status:
                st.write("Opening and regenerating the model through CREOSON")
                result = AuditEngine(config).run()
                st.write("Writing interactive HTML, JSON, and CSV reports")
                outputs = write_reports(result, config, config.export_directory)
                status.update(label="Audit complete", state="complete")
            st.session_state["audit_result"] = result
            st.session_state["audit_config"] = config
            st.session_state["audit_outputs"] = outputs
        except Exception as exc:
            st.exception(exc)

result = st.session_state.get("audit_result")
if result:
    summary = result.summary()
    cols = st.columns(5)
    cols[0].metric("Bodies", len(result.bodies))
    cols[1].metric("Errors", summary["error"])
    cols[2].metric("Warnings", summary["warning"])
    cols[3].metric("Passes", summary["pass"])
    cols[4].metric("Units", result.length_units)

    st.plotly_chart(build_figure(result), use_container_width=True)
    findings_tab, bodies_tab, warnings_tab, downloads_tab = st.tabs(["Findings", "Bodies", "Warnings", "Reports"])
    with findings_tab:
        rows = []
        for item in result.findings_dicts():
            details = item.pop("details")
            item.update(
                {
                    "method": details.get("method"),
                    "patch_a": details.get("patch_a"),
                    "patch_b": details.get("patch_b"),
                    "allowed_maximum": details.get("allowed_maximum"),
                    "raw_distance": details.get("raw_distance"),
                }
            )
            rows.append(item)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    with bodies_tab:
        st.dataframe(pd.DataFrame(result.metadata["body_summary"]), use_container_width=True, hide_index=True)
    with warnings_tab:
        if result.warnings:
            for warning in result.warnings:
                st.warning(warning)
        else:
            st.success("No analysis warnings were generated.")
    with downloads_tab:
        output_paths = st.session_state["audit_outputs"]
        archive = io.BytesIO()
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
            for path in output_paths.values():
                bundle.write(path, arcname=Path(path).name)
        st.download_button(
            "Download report bundle", archive.getvalue(),
            file_name=f"{Path(result.model_name).stem}_audit_reports.zip", mime="application/zip",
        )
