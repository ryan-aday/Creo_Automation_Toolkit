"""
Logic-only tests for the hierarchy rollup.

These tests do not connect to Creo. They import the report module with a tiny
creopyson stub so the tree/grouping logic can be checked independently.
"""

from __future__ import annotations

import importlib.util
import math
import sys
import types
from pathlib import Path


def load_module():
    stub = types.ModuleType("creopyson")
    class DummyClient:
        pass
    stub.Client = DummyClient
    sys.modules["creopyson"] = stub

    module_path = Path(__file__).resolve().parents[1] / "creo_assembly_mass_report.py"
    spec = importlib.util.spec_from_file_location("report_module", module_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules["report_module"] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def build_fixture(m):
    # Same bolt twice in SUB_A and twice in SUB_C.
    bom = {
        "file": "TOP.asm",
        "children": [
            {
                "file": "SUB_A.asm",
                "seq_path": "root.1",
                "children": [
                    {"file": "NAS_BOLT.prt", "seq_path": "root.1.1"},
                    {"file": "NAS_BOLT.prt", "seq_path": "root.1.2"},
                    {"file": "PLATE.prt", "seq_path": "root.1.3"},
                ],
            },
            {
                "file": "SUB_C.asm",
                "seq_path": "root.3",
                "children": [
                    {"file": "NAS_BOLT.prt", "seq_path": "root.3.1"},
                    {"file": "NAS_BOLT.prt", "seq_path": "root.3.2"},
                ],
            },
        ],
    }
    root = m.parse_bom_tree(bom, "TOP.asm")
    md = {}
    for model, mass in {
        "NAS_BOLT.prt": 0.1,
        "PLATE.prt": 1.0,
    }.items():
        md[model] = m.ModelMetadata(
            model=model,
            model_type="PART",
            native_mass_unit="kg",
            native_mass=mass,
        )
    for model in ("SUB_A.asm", "SUB_C.asm"):
        md[model] = m.ModelMetadata(model=model, model_type="ASSEMBLY")

    categories = m.collections.OrderedDict([("FASTENER", ["NAS"])])
    m.assign_occurrence_properties(root, md, categories)
    return root, md


def test_parent_local_deduplication():
    m = load_module()
    root, md = build_fixture(m)
    rows = m.group_hierarchy_rows(root, md)

    bolt_rows = [r for r in rows if r.model == "NAS_BOLT.prt"]
    assert len(bolt_rows) == 2
    assert sorted(r.qty for r in bolt_rows) == [2, 2]
    assert sorted(r.immediate_assembly for r in bolt_rows) == ["SUB_A.asm", "SUB_C.asm"]
    assert sorted(r.top_qty for r in bolt_rows) == [4, 4]


def test_unique_rollup():
    m = load_module()
    root, md = build_fixture(m)
    rows = m.group_unique_rows(root, md)
    bolt = next(r for r in rows if r.model == "NAS_BOLT.prt")
    assert bolt.qty == 4
    assert len(bolt.indices) == 4
    assert math.isclose(bolt.total_mass_kg, 0.4)


def test_recursive_assembly_mass():
    m = load_module()
    root, md = build_fixture(m)
    sub_a = next(o for o in root.children if o.model == "SUB_A.asm")
    sub_c = next(o for o in root.children if o.model == "SUB_C.asm")
    assert math.isclose(sub_a.mass_kg, 1.2)
    assert math.isclose(sub_c.mass_kg, 0.2)
    assert math.isclose(root.mass_kg, 1.4)


def test_fastener_separation():
    m = load_module()
    root, md = build_fixture(m)
    breakdown = dict(m.top_level_mass_breakdown(root))
    assert math.isclose(breakdown["SUB_A.asm"], 1.0)
    assert math.isclose(breakdown["SUB_C.asm"], 0.0)
    assert math.isclose(breakdown["FASTENER (all leaf parts)"], 0.4)
