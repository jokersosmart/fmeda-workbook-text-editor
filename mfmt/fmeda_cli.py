from __future__ import annotations

import argparse
from pathlib import Path

from .spreadsheet.fmeda_workspace import FmedaCoreWorkspaceBuilder


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a source-safe, reviewable FMEDA text workspace."
    )
    parser.add_argument("input", type=Path, help="Source .xlsx workbook")
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=Path("fmeda_workspace"),
        help="Workspace output directory",
    )
    parser.add_argument(
        "--adapter",
        choices=("core", "editor"),
        default="core",
        help="Optional integration adapter. Default: core-only output.",
    )
    args = parser.parse_args()

    include_editor = args.adapter == "editor"
    manifest = FmedaCoreWorkspaceBuilder(
        args.input,
        args.output_dir,
        include_editor=include_editor,
    ).build()
    print(
        f"OK: {args.output_dir} "
        f"({manifest['sheet_count']} sheets, {manifest['formula_count']} formulas)"
    )
    print(f"Core workspace: {args.output_dir / 'normalized'}")
    print(f"Readable entrypoint: {args.output_dir / 'readable' / 'index.md'}")
    print(f"Derived workbook: {args.output_dir / manifest['derived_file']}")
    if manifest.get("editor"):
        print(f"Editor adapter: {args.output_dir / manifest['editor']['root']}")
    else:
        print("Editor adapter: disabled (core-only mode)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
