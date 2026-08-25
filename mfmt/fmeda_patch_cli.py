from __future__ import annotations

import argparse
import json
from pathlib import Path

from .spreadsheet.fmeda_patch import FmedaPatchApplier


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply a formula-safe patch to a derived FMEDA workbook."
    )
    parser.add_argument("workspace", type=Path, help="FMEDA workspace directory")
    parser.add_argument("patch", type=Path, help="JSON patch manifest")
    args = parser.parse_args()
    patch = json.loads(args.patch.read_text(encoding="utf-8"))
    result = FmedaPatchApplier(args.workspace).apply(patch)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
