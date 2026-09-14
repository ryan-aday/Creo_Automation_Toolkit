from __future__ import annotations

import csv
import html
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

import plotly.io as pio

from .models import AuditConfig, AuditResult
from .visualization import build_figure


def _json_default(value: Any):
    if hasattr(value, "tolist"):
        return value.tolist()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, set):
        return sorted(value)
    return str(value)


def write_reports(result: AuditResult, config: AuditConfig, directory: str | Path) -> dict[str, Path]:
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    stem = Path(result.model_name).stem
    json_path = directory / f"{stem}_audit.json"
    csv_path = directory / f"{stem}_findings.csv"
    html_path = directory / f"{stem}_audit.html"

    payload = {
        "model_name": result.model_name,
        "length_units": result.length_units,
        "summary": result.summary(),
        "warnings": result.warnings,
        "metadata": result.metadata,
        "config": asdict(config),
        "findings": result.findings_dicts(),
    }
    json_path.write_text(json.dumps(payload, indent=2, default=_json_default), encoding="utf-8")

    fieldnames = ["check", "severity", "message", "body_a", "body_b", "value", "limit", "units", "details"]
    with csv_path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        for row in result.findings_dicts():
            row["details"] = json.dumps(row["details"], default=_json_default)
            writer.writerow(row)

    figure_html = pio.to_html(build_figure(result), include_plotlyjs=True, full_html=False)
    rows = "".join(
        "<tr>"
        f"<td>{html.escape(item.severity)}</td><td>{html.escape(item.check)}</td>"
        f"<td>{html.escape(item.body_a or '')}</td><td>{html.escape(item.body_b or '')}</td>"
        f"<td>{'' if item.value is None else f'{item.value:.8g}'}</td>"
        f"<td>{html.escape(item.units)}</td><td>{html.escape(item.message)}</td>"
        "</tr>"
        for item in result.findings
    )
    warning_html = "".join(f"<li>{html.escape(warning)}</li>" for warning in result.warnings)
    summary = result.summary()
    html_path.write_text(
        f"""<!doctype html><html><head><meta charset="utf-8"><title>{html.escape(stem)} audit</title>
<style>body{{font:14px system-ui;margin:2rem;color:#17212b}} .cards{{display:flex;gap:1rem}} .card{{padding:1rem;border:1px solid #d7e0e8;border-radius:8px}} table{{border-collapse:collapse;width:100%}} th,td{{padding:.5rem;border-bottom:1px solid #dde5ec;text-align:left}} th{{background:#edf3f7;position:sticky;top:0}} .note{{background:#fff7e6;padding:1rem;border-left:4px solid #e6a700}}</style></head><body>
<h1>Creo geometry audit — {html.escape(result.model_name)}</h1>
<div class="cards"><div class="card">Errors: {summary['error']}</div><div class="card">Warnings: {summary['warning']}</div><div class="card">Passes: {summary['pass']}</div><div class="card">Bodies: {len(result.bodies)}</div></div>
<p class="note">Results are based on a tessellated STEP snapshot. Mesh clearance and patch areas are screening measurements, not replacements for Creo's B-rep measurement or formal variation analysis.</p>
{figure_html}<h2>Findings</h2><table><thead><tr><th>Severity</th><th>Check</th><th>Body A</th><th>Body B</th><th>Value</th><th>Units</th><th>Message</th></tr></thead><tbody>{rows}</tbody></table>
<h2>Warnings</h2><ul>{warning_html}</ul></body></html>""",
        encoding="utf-8",
    )
    return {"json": json_path, "csv": csv_path, "html": html_path}

