from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import numpy as np
import trimesh

from .models import BodyRecord
from .units import canonical_length_unit


def _clean_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("_")
    return value or "body"


def _source_model(name: str, component_names: list[str]) -> str | None:
    probe = name.lower()
    ranked: list[tuple[int, str]] = []
    for component in component_names:
        stem = Path(component).stem.lower()
        if stem and (stem in probe or probe in stem):
            ranked.append((len(stem), component))
    return max(ranked, default=(0, None))[1]


def load_step_bodies(
    step_path: str | Path,
    target_units: str,
    components: list[dict[str, Any]] | None = None,
) -> tuple[list[BodyRecord], list[str]]:
    """Load a STEP scene while retaining occurrence transforms and body identity."""
    warnings: list[str] = []
    try:
        scene = trimesh.load_scene(str(step_path), file_type="step", process=True)
    except BaseException as exc:
        raise RuntimeError(
            "STEP tessellation failed. Install the pinned Trimesh/Cascadio dependencies and "
            f"confirm the exported STEP is valid. Original error: {exc}"
        ) from exc

    target = canonical_length_unit(target_units)
    if scene.units:
        try:
            converted = scene.convert_units(target, guess=False)
            if converted is not None:
                scene = converted
        except Exception as exc:
            warnings.append(
                f"STEP units '{scene.units}' could not be converted to '{target}': {exc}. "
                "Thresholds may not be in the same unit system."
            )
    else:
        warnings.append(
            f"STEP scene did not declare units; coordinates were assumed to be {target}."
        )

    component_names = []
    for record in components or []:
        for key in ("file", "name", "generic"):
            if record.get(key):
                component_names.append(str(record[key]))
                break

    bodies: list[BodyRecord] = []
    nodes = list(scene.graph.nodes_geometry)
    if not nodes:
        # A direct Trimesh can appear as one unnamed geometry.
        nodes = list(scene.geometry.keys())

    serial = 0
    for node_name in nodes:
        try:
            transform, geometry_name = scene.graph[node_name]
            geometry = scene.geometry[geometry_name]
        except Exception:
            geometry_name = node_name
            geometry = scene.geometry[node_name]
            transform = np.eye(4)
        if not isinstance(geometry, trimesh.Trimesh):
            continue
        world = geometry.copy()
        world.apply_transform(np.asarray(transform, dtype=float))
        pieces = list(world.split(only_watertight=False)) or [world]
        for piece_index, piece in enumerate(pieces, start=1):
            serial += 1
            base = _clean_name(str(node_name))
            name = base if len(pieces) == 1 else f"{base}__solid_{piece_index}"
            bodies.append(
                BodyRecord(
                    body_id=f"B{serial:04d}",
                    name=name,
                    geometry_name=str(geometry_name),
                    node_name=str(node_name),
                    source_model=_source_model(f"{node_name} {geometry_name}", component_names),
                    mesh=piece,
                )
            )
    if not bodies:
        raise RuntimeError("No tessellated solid bodies were found in the exported STEP scene.")
    return bodies, warnings
