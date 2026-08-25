from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_large import LargeFmedaValidator


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run resumable, metadata-aware validation for large FMEDA revisions."
    )
    parser.add_argument("base", type=Path, help="Base workbook revision")
    parser.add_argument("target", type=Path, help="Target workbook revision")
    parser.add_argument("--output-dir", "-o", type=Path, required=True)
    parser.add_argument("--patch-manifest", type=Path)
    parser.add_argument(
        "--max-sheets",
        type=int,
        help="Process at most this many not-yet-completed sheets, then exit with INCOMPLETE",
    )
    args = parser.parse_args()
    patch = {}
    if args.patch_manifest:
        patch = json.loads(args.patch_manifest.read_text(encoding="utf-8"))
    report = LargeFmedaValidator(
        args.base,
        args.target,
        args.output_dir,
        patch,
    ).run(max_sheets=args.max_sheets)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] in {"PASS", "PASS_WITH_WARNINGS"} else 2


if __name__ == "__main__":
    raise SystemExit(main())
