from __future__ import annotations

import numpy as np

from .models import BodyRecord


def thermally_expand(body: BodyRecord, delta_temperature: float, anchor_mode: str) -> BodyRecord:
    if body.cte_per_k is None:
        raise ValueError(f"No CTE is available for {body.name}")
    factor = 1.0 + body.cte_per_k * delta_temperature
    if factor <= 0:
        raise ValueError(f"Thermal scale factor is non-positive for {body.name}")
    mesh = body.mesh.copy()
    anchor = np.asarray(mesh.centroid if anchor_mode == "centroid" else [0.0, 0.0, 0.0])
    mesh.vertices = anchor + factor * (np.asarray(mesh.vertices) - anchor)
    return BodyRecord(
        body_id=body.body_id,
        name=body.name,
        geometry_name=body.geometry_name,
        node_name=body.node_name,
        source_model=body.source_model,
        mesh=mesh,
        material=body.material,
        cte_per_k=body.cte_per_k,
    )

