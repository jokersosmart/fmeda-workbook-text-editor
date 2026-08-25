from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_acceptance import FmedaAcceptanceProfile


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Classify external FMEDA recalculation results as accepted, review_required, or blocked."
    )
    parser.add_argument("base", type=Path, help="Original or baseline workbook")
    parser.add_argument("target", type=Path, help="Recalculated workbook")
    parser.add_argument("--recalc-report", type=Path, required=True, help="JSON report from fmeda-recalculate")
    parser.add_argument("--decision-manifest", type=Path, help="Optional reviewer decision JSON")
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("acceptance-report.md"),
        help="Markdown output path",
    )
    parser.add_argument("--json-output", type=Path, help="Optional JSON output path")
    args = parser.parse_args()

    recalc_report = json.loads(args.recalc_report.read_text(encoding="utf-8"))
    decision_manifest = None
    if args.decision_manifest:
        decision_manifest = json.loads(args.decision_manifest.read_text(encoding="utf-8"))
    report = FmedaAcceptanceProfile().write_reports(
        args.base,
        args.target,
        recalc_report,
        args.output,
        args.json_output,
        decision_manifest,
    )
    print(f"{report['status']}: {args.output}")
    print(
        "accepted={accepted} review_required={review_required} blocked={blocked}".format(
            **report["counts"]
        )
    )
    return 2 if report["status"] == "blocked" else 0


if __name__ == "__main__":
    raise SystemExit(main())
