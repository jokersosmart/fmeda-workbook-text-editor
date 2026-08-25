from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_diff import FmedaRevisionValidator


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate traceable differences between two FMEDA revisions."
    )
    parser.add_argument("base", type=Path, help="Base workbook revision")
    parser.add_argument("target", type=Path, help="Target workbook revision")
    parser.add_argument("--patch-manifest", type=Path, help="Applied patch manifest JSON")
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("validation-report.md"),
        help="Markdown report output path",
    )
    parser.add_argument(
        "--json-output",
        type=Path,
        help="Optional JSON report output path",
    )
    args = parser.parse_args()
    patch_manifest = {}
    if args.patch_manifest:
        patch_manifest = json.loads(args.patch_manifest.read_text(encoding="utf-8"))
    validator = FmedaRevisionValidator(args.base, args.target, patch_manifest)
    report = validator.write_report(args.output)
    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    print(f"{report['status']}: {args.output}")
    print(
        f"allowed={report['allowed_input_changes']} "
        f"blocking={report['blocking_change_count']} "
        f"warnings={report['warning_count']}"
    )
    return 0 if report["status"] != "FAIL" else 2


if __name__ == "__main__":
    raise SystemExit(main())
