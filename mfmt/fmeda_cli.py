from __future__ import annotations

import argparse
from pathlib import Path

from .spreadsheet.fmeda_workspace import FmedaWorkspaceBuilder


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a source-safe, reviewable FMEDA workbook workspace."
    )
    parser.add_argument("input", type=Path, help="Source .xlsx workbook")
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=Path("fmeda_workspace"),
        help="Workspace output directory",
    )
    args = parser.parse_args()
    manifest = FmedaWorkspaceBuilder(args.input, args.output_dir).build()
    print(
        f"OK: {args.output_dir} "
        f"({manifest['sheet_count']} sheets, {manifest['formula_count']} formulas)"
    )
    print(f"Editor workspace: {args.output_dir / 'editor'}")
    print(f"Derived workbook: {args.output_dir / manifest['derived_file']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
