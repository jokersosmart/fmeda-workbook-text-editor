from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_acceptance import FmedaAcceptanceProfile
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
    parser.add_argument(
        "--recalc-report",
        type=Path,
        help="Optional JSON report from fmeda-recalculate; enables external acceptance profiling",
    )
    parser.add_argument(
        "--decision-manifest",
        type=Path,
        help="Optional reviewer decision manifest for external acceptance",
    )
    parser.add_argument(
        "--acceptance-output",
        type=Path,
        help="Optional Markdown output for the acceptance profile",
    )
    parser.add_argument(
        "--acceptance-json",
        type=Path,
        help="Optional JSON output for the acceptance profile",
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

    acceptance_report = None
    if args.recalc_report:
        acceptance_output = args.acceptance_output or args.output.with_name(
            f"{args.output.stem}.acceptance.md"
        )
        acceptance_json = args.acceptance_json
        recalc_report = json.loads(args.recalc_report.read_text(encoding="utf-8"))
        decision_manifest = None
        if args.decision_manifest:
            decision_manifest = json.loads(args.decision_manifest.read_text(encoding="utf-8"))
        acceptance_report = FmedaAcceptanceProfile().write_reports(
            args.base,
            args.target,
            recalc_report,
            acceptance_output,
            acceptance_json,
            decision_manifest,
        )

    print(f"{report['status']}: {args.output}")
    print(
        f"allowed={report['allowed_input_changes']} "
        f"blocking={report['blocking_change_count']} "
        f"warnings={report['warning_count']}"
    )
    if acceptance_report:
        print(
            f"acceptance={acceptance_report['status']} "
            f"accepted={acceptance_report['counts']['accepted']} "
            f"review_required={acceptance_report['counts']['review_required']} "
            f"blocked={acceptance_report['counts']['blocked']}"
        )
    return 2 if report["status"] == "FAIL" or (
        acceptance_report is not None and acceptance_report["status"] == "blocked"
    ) else 0


if __name__ == "__main__":
    raise SystemExit(main())
