from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from .engine import AuditEngine
from .models import AuditConfig
from .reporting import write_reports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Audit a Creo part or assembly through CREOSON.")
    parser.add_argument("model", nargs="?", help="Absolute or working-directory-relative .prt/.asm path")
    parser.add_argument("--config", type=Path, help="JSON configuration file")
    parser.add_argument("--host", default="localhost", help="CREOSON host")
    parser.add_argument("--port", type=int, default=9056, help="CREOSON port")
    parser.add_argument("--creo-version", type=int, help="Creo major version")
    parser.add_argument("--output", type=Path, default=Path("audit_output"), help="Output directory")
    parser.add_argument("--minimum-gap", type=float, default=0.0, help="Global minimum gap in model units")
    parser.add_argument("--minimum-face-area", type=float, default=0.0, help="Patch-area threshold in model units²")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return parser


def main() -> int:
    args = build_parser().parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level), format="%(levelname)s %(message)s")
    if args.config:
        raw = json.loads(args.config.read_text(encoding="utf-8"))
        if args.model:
            raw["model_path"] = args.model
        config = AuditConfig.from_dict(raw)
    else:
        if not args.model:
            raise SystemExit("MODEL is required unless --config supplies model_path")
        config = AuditConfig(
            model_path=args.model,
            creoson_host=args.host,
            creoson_port=args.port,
            creo_version=args.creo_version,
            export_directory=str(args.output),
            minimum_global_gap=args.minimum_gap,
            minimum_face_area=args.minimum_face_area,
        )

    result = AuditEngine(config).run()
    outputs = write_reports(result, config, config.export_directory)
    counts = result.summary()
    print(
        f"Audit complete: {counts['error']} errors, {counts['warning']} warnings, "
        f"{counts['pass']} passes across {len(result.bodies)} bodies."
    )
    for kind, path in outputs.items():
        print(f"{kind}: {path.resolve()}")
    return 2 if counts["error"] else 0


if __name__ == "__main__":
    raise SystemExit(main())

