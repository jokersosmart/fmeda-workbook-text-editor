from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_recalc import LibreOfficeRecalculator


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Recalculate an FMEDA workbook using LibreOffice Calc without changing the source."
    )
    parser.add_argument("source", type=Path, help="Source .xlsx or .xlsm workbook")
    parser.add_argument("output", type=Path, help="Recalculated output workbook")
    parser.add_argument("--timeout", type=int, default=900, help="Timeout in seconds")
    parser.add_argument("--report", type=Path, help="Optional JSON recalculation report")
    args = parser.parse_args()
    report = LibreOfficeRecalculator(
        args.source,
        args.output,
        timeout_seconds=args.timeout,
    ).run()
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
