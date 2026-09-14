from __future__ import annotations

import numpy as np
import plotly.graph_objects as go

from .models import AuditResult


def _display_faces(mesh, max_faces: int):
    faces = np.asarray(mesh.faces)
    if len(faces) <= max_faces:
        return np.asarray(mesh.vertices), faces
    selected = np.linspace(0, len(faces) - 1, max_faces, dtype=int)
    return np.asarray(mesh.vertices), faces[selected]


def build_figure(result: AuditResult, max_faces_per_body: int = 25_000) -> go.Figure:
    colliding = set(result.metadata.get("colliding_body_ids", []))
    figure = go.Figure()
    for body in result.bodies:
        vertices, faces = _display_faces(body.mesh, max_faces_per_body)
        is_collision = body.body_id in colliding
        figure.add_trace(
            go.Mesh3d(
                x=vertices[:, 0], y=vertices[:, 1], z=vertices[:, 2],
                i=faces[:, 0], j=faces[:, 1], k=faces[:, 2],
                name=body.name,
                color="#e53935" if is_collision else "#8fa7bd",
                opacity=0.96 if is_collision else 0.22,
                flatshading=False,
                hovertemplate=(
                    f"<b>{body.name}</b><br>Model: {body.source_model or 'unmapped'}"
                    f"<br>Material: {body.material or 'unset'}<extra></extra>"
                ),
                showscale=False,
            )
        )

    for index, finding in enumerate(result.findings):
        if not finding.check.startswith(("minimum_gap", "gap_rule")):
            continue
        point_a = finding.details.get("point_a")
        point_b = finding.details.get("point_b")
        if point_a is None or point_b is None:
            continue
        color = "#e53935" if finding.severity == "error" else "#2e7d32"
        figure.add_trace(
            go.Scatter3d(
                x=[point_a[0], point_b[0]], y=[point_a[1], point_b[1]], z=[point_a[2], point_b[2]],
                mode="lines+markers",
                line={"color": color, "width": 7}, marker={"color": color, "size": 4},
                name=f"Gap {index + 1}",
                hovertemplate=(
                    f"{finding.body_a} ↔ {finding.body_b}<br>"
                    f"Clearance: {finding.value:.6g} {finding.units}<extra></extra>"
                ),
            )
        )

    figure.update_layout(
        template="plotly_white",
        margin={"l": 0, "r": 0, "t": 42, "b": 0},
        title=f"{result.model_name} — geometry audit",
        legend={"orientation": "h", "y": -0.08},
        scene={
            "aspectmode": "data",
            "xaxis_title": f"X ({result.length_units})",
            "yaxis_title": f"Y ({result.length_units})",
            "zaxis_title": f"Z ({result.length_units})",
        },
    )
    return figure

