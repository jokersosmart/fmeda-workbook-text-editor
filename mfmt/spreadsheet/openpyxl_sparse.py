"""Bounded sparse-cell access for normal-mode openpyxl worksheets."""

from __future__ import annotations

from typing import Any


def populated_cells(worksheet: Any) -> list[Any]:
    """Return value/formula cells without expanding the declared rectangle.

    Enterprise templates often contain formatting at Excel's last row.  The
    public ``iter_rows`` API expands that formatting-only dimension and can
    create tens of millions of empty cell objects.  Normal-mode openpyxl keeps
    instantiated cells in ``_cells``; isolate that private compatibility seam
    here and retain a public-API fallback for worksheet-like implementations.
    """
    cell_store = getattr(worksheet, "_cells", None)
    if isinstance(cell_store, dict):
        return [
            cell
            for _coordinate, cell in sorted(cell_store.items())
            if getattr(cell, "value", None) is not None
        ]

    return [
        cell
        for row in worksheet.iter_rows()
        for cell in row
        if getattr(cell, "value", None) is not None
    ]


def effective_dimensions(worksheet: Any, cells: list[Any]) -> dict[str, int]:
    """Return content bounds while preserving non-empty merged-range extents."""
    if not cells:
        return {"max_row": 1, "max_col": 1}

    max_row = max(cell.row for cell in cells)
    max_col = max(cell.column for cell in cells)
    coordinates = {(cell.row, cell.column) for cell in cells}
    for merged in worksheet.merged_cells.ranges:
        if (merged.min_row, merged.min_col) in coordinates:
            max_row = max(max_row, merged.max_row)
            max_col = max(max_col, merged.max_col)
    return {"max_row": max_row, "max_col": max_col}
