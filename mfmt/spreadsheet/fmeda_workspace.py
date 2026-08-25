"""Build a source-safe, reviewable FMEDA workbook workspace.

The first slice deliberately uses the existing rich Excel converter as the
low-level reader.  This module owns the workspace/provenance contract and
never writes to the input workbook.
"""

from __future__ import annotations

import csv
import hashlib
import json
import re
import shutil
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from .xlsx_to_json import ExcelToJsonConverter


SCHEMA_VERSION = "workbook-v2"
PROFILE = "spreadsheet-fmeda"
EDITOR_SIDECAR_SCHEMA = "fmeda-editor-sidecar-v1"
RELATIONS_SCHEMA = "editor-relations-v0.2"
MAX_FORMULA_BLOCKS_PER_SHEET = 2_000

_CELL_REF_RE = re.compile(r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")
_EXTERNAL_REF_RE = re.compile(
    r"(?:\[(?P<workbook>[^\]]+)\])?"
    r"(?:'(?P<quoted_sheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_. -]*))!"
    r"(?P<cell>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"
)
_FORMULA_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9_.]*)\s*\(")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_stem(path: Path) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._") or "workbook"


def _safe_sheet_name(name: str, index: int) -> str:
    cleaned = re.sub(r"[^\w\s-]", "", name, flags=re.UNICODE).strip().replace(" ", "_")
    return f"{index:02d}_{cleaned or 'Sheet'}"


def _json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def _is_error(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("#")


def _strip_formula_strings(formula: str) -> str:
    return _FORMULA_STRING_RE.sub(" ", formula)


def _normalize_ref(value: str) -> str:
    return value.replace("$", "")


def _formula_references(formula: str, current_sheet: str) -> list[dict[str, str]]:
    """Return explicit external/cross-sheet and implicit same-sheet refs."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return []

    body = _strip_formula_strings(formula)
    refs: list[dict[str, str]] = []
    covered: list[tuple[int, int]] = []

    for match in _EXTERNAL_REF_RE.finditer(body):
        workbook = match.group("workbook") or ""
        sheet = match.group("quoted_sheet") or match.group("sheet") or ""
        cell_range = _normalize_ref(match.group("cell"))
        kind = "external_reference" if workbook else (
            "same_sheet" if sheet == current_sheet else "cross_sheet"
        )
        refs.append(
            {
                "reference_kind": kind,
                "workbook": workbook,
                "sheet": sheet,
                "range": cell_range,
            }
        )
        covered.append((match.start(), match.end()))

    def is_covered(start: int, end: int) -> bool:
        return any(start >= left and end <= right for left, right in covered)

    for match in _CELL_REF_RE.finditer(body):
        if is_covered(match.start(), match.end()):
            continue
        # A cell immediately following ! was already handled as a qualified ref.
        if match.start() > 0 and body[match.start() - 1] == "!":
            continue
        refs.append(
            {
                "reference_kind": "same_sheet",
                "workbook": "",
                "sheet": current_sheet,
                "range": _normalize_ref(match.group(0)),
            }
        )

    # Keep stable order while removing duplicates from overlapping regex matches.
    unique: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for ref in refs:
        key = (
            ref["reference_kind"],
            ref["workbook"],
            ref["sheet"],
            ref["range"],
        )
        if key not in seen:
            seen.add(key)
            unique.append(ref)
    return unique


def _function_names(formula: str) -> list[str]:
    if not isinstance(formula, str):
        return []
    return sorted({match.group(1).upper() for match in _FUNCTION_RE.finditer(formula)})


def _cell_status(cell: dict[str, Any]) -> tuple[str, str]:
    formula = cell.get("formula")
    value = cell.get("value")
    if formula:
        if _is_error(value) or "#DIV/0!" in str(formula) or "#VALUE!" in str(formula):
            return "error", "formula_result"
        if value is None:
            return "not_calculated", "formula_result"
        return "cached_value", "formula_result"
    if _is_error(value):
        return "error", "source_error"
    if value is None:
        return "empty", "empty"
    return "input_or_constant", "input_or_constant"


def _relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


class FmedaWorkspaceBuilder:
    """Build a source-safe FMEDA core workspace from one XLSX file.

    Editor output is optional and emitted by ``FmedaEditorAdapter`` so the
    core can run without importing or requiring the Editor integration.
    ``FmedaWorkspaceBuilder`` remains the compatibility name for the core
    builder.
    """

    def __init__(self, source: str | Path, workspace: str | Path, *, include_editor: bool = False):
        self.source = Path(source).expanduser().resolve()
        self.workspace = Path(workspace).expanduser().resolve()
        self.include_editor = include_editor
        if not self.source.is_file():
            raise FileNotFoundError(f"FMEDA source workbook not found: {self.source}")
        if self.source.suffix.lower() != ".xlsx":
            raise ValueError("FMEDA workspace currently supports .xlsx inputs only")
        if self.source == self.workspace:
            raise ValueError("workspace must be different from the source workbook")

    def build(self) -> dict[str, Any]:
        self._make_dirs()
        source_copy = self.workspace / "source" / self.source.name
        if source_copy.resolve() != self.source.resolve():
            shutil.copy2(self.source, source_copy)
        source_hash = _sha256(self.source)

        derived_name = f"{_safe_stem(self.source)}.rev-001{self.source.suffix.lower()}"
        derived_copy = self.workspace / "derived" / derived_name
        shutil.copy2(self.source, derived_copy)

        rich_dir = self.workspace / "normalized" / "sheets"
        converter = ExcelToJsonConverter(
            str(self.source),
            str(rich_dir),
            skip_ocr=True,
            extract_images=False,
        )
        converter.convert()

        legacy_manifest = self._read_json(rich_dir / "workbook.json")
        sheet_models: list[dict[str, Any]] = []
        formula_rows: list[dict[str, Any]] = []
        edge_rows: list[dict[str, Any]] = []
        review_items: list[dict[str, Any]] = []
        sheet_entries: list[dict[str, Any]] = []

        for sheet_entry in legacy_manifest.get("sheets", []):
            index = int(sheet_entry.get("index") or len(sheet_models) + 1)
            name = str(sheet_entry.get("name") or f"Sheet{index}")
            legacy_path = rich_dir / Path(str(sheet_entry.get("json_file", ""))).name
            legacy_sheet = self._read_json(legacy_path)
            model, sheet_formula_rows, sheet_edges, sheet_reviews = self._normalize_sheet(
                legacy_sheet, source_hash
            )
            model["sheet_meta"]["index"] = index
            model["sheet_meta"]["name"] = name
            sheet_json_path = rich_dir / f"{_safe_sheet_name(name, index)}.json"
            sheet_json_path.write_text(
                json.dumps(model, ensure_ascii=False, indent=2, default=str) + "\n",
                encoding="utf-8",
            )
            sheet_models.append(model)
            formula_rows.extend(sheet_formula_rows)
            edge_rows.extend(sheet_edges)
            review_items.extend(sheet_reviews)
            md_name = f"{_safe_sheet_name(name, index)}.md"
            sheet_manifest = {
                "index": index,
                "name": name,
                "json_file": f"normalized/sheets/{sheet_json_path.name}",
                "formula_count": len(sheet_formula_rows),
                "review_count": len(sheet_reviews),
            }
            if self.include_editor:
                sheet_manifest["md_file"] = f"editor/sheets/{md_name}"
            sheet_entries.append(sheet_manifest)

        workbook_model = {
            "schema_version": SCHEMA_VERSION,
            "profile": PROFILE,
            "source": {
                "filename": self.source.name,
                "source_file": _relative(source_copy, self.workspace),
                "sha256": source_hash,
                "calculation_mode": "source_cached_values",
            },
            "workbook_meta": {
                "sheet_order": [entry["name"] for entry in sheet_entries],
                "external_links": sorted(
                    {item["details"].get("workbook", "") for item in review_items if item["kind"] == "external_reference"}
                    - {""}
                ),
            },
            "sheets": sheet_entries,
            "risk_flags": sorted({item["kind"] for item in review_items}),
            "validation_summary": {
                "sheet_count": len(sheet_entries),
                "formula_count": len(formula_rows),
                "dependency_edge_count": len(edge_rows),
                "review_item_count": len(review_items),
            },
        }
        workbook_path = self.workspace / "normalized" / "Step03_workbook.json"
        workbook_path.write_text(
            json.dumps(workbook_model, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        self._write_formula_catalog(formula_rows)
        self._write_dependency_edges(edge_rows)
        self._write_review_items(review_items)
        self._write_summary(workbook_model, formula_rows, edge_rows, review_items)
        editor_manifest = None
        if self.include_editor:
            from .fmeda_editor_adapter import FmedaEditorAdapter

            editor_manifest = FmedaEditorAdapter().emit(
                self.workspace,
                workbook_model,
                sheet_entries,
                formula_rows,
                review_items,
                source_hash,
            )

        exported_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        manifest = {
            "schema_version": SCHEMA_VERSION,
            "profile": PROFILE,
            "source_file": _relative(source_copy, self.workspace),
            "source_sha256": source_hash,
            "derived_file": _relative(derived_copy, self.workspace),
            "source_original_path": str(self.source),
            "created_at": exported_at,
            "sheet_count": len(sheet_entries),
            "formula_count": len(formula_rows),
            "dependency_edge_count": len(edge_rows),
            "review_item_count": len(review_items),
            "calculation_mode": "source_cached_values",
            "editor": editor_manifest,
            "sheets": sheet_entries,
        }
        (self.workspace / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        report = self._build_import_report(manifest, review_items)
        (self.workspace / "reports" / "import-report.md").write_text(report, encoding="utf-8")
        return manifest

    def _make_dirs(self) -> None:
        directories = [
            "source",
            "derived",
            "normalized/sheets",
            "reports",
            "mappings",
        ]
        if self.include_editor:
            directories.append("editor/sheets")
        for relative in directories:
            (self.workspace / relative).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any]:
        return json.loads(path.read_text(encoding="utf-8"))

    def _normalize_sheet(
        self, legacy_sheet: dict[str, Any], source_hash: str
    ) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
        meta = dict(legacy_sheet.get("sheet_meta") or {})
        name = str(meta.get("name") or "Sheet")
        cells: dict[str, Any] = {}
        formula_rows: list[dict[str, Any]] = []
        edge_rows: list[dict[str, Any]] = []
        review_items: list[dict[str, Any]] = []

        for key, raw_cell in (legacy_sheet.get("cells") or {}).items():
            cell = dict(raw_cell)
            address = str(cell.get("address") or key)
            status, source_kind = _cell_status(cell)
            formula = cell.get("formula")
            refs = _formula_references(str(formula), name) if formula else []
            cell["calculation_status"] = status
            cell["source_kind"] = source_kind
            cell["cached_value"] = _json_value(cell.get("value")) if formula else None
            cell["display_value"] = _json_value(cell.get("value"))
            cell["formula_raw"] = formula if formula else None
            cell["formula_refs"] = refs
            cell["provenance"] = {
                "source": "xlsx",
                "source_sha256": source_hash,
                "source_cell": f"{name}!{address}",
            }
            cells[key] = cell

            if formula:
                formula_id = f"FORMULA-{_safe_stem(Path(name))}-{address}"
                row = {
                    "formula_id": formula_id,
                    "sheet": name,
                    "cell": address,
                    "source_cell": f"{name}!{address}",
                    "formula_raw": str(formula),
                    "cached_value": _json_value(cell.get("value")),
                    "status": status,
                    "function_names": ";".join(_function_names(str(formula))),
                    "dependencies": json.dumps(refs, ensure_ascii=False, separators=(",", ":")),
                }
                formula_rows.append(row)
                cell["formula_id"] = formula_id
                for ref in refs:
                    target = f"{ref['sheet']}!{ref['range']}"
                    edge_rows.append(
                        {
                            "source": f"{name}!{address}",
                            "reference_kind": ref["reference_kind"],
                            "target": target,
                            "target_sheet": ref["sheet"],
                            "target_range": ref["range"],
                            "external_workbook": ref["workbook"],
                            "status": "unresolved" if ref["reference_kind"] == "external_reference" else "parsed",
                        }
                    )
                    if ref["reference_kind"] == "external_reference":
                        review_items.append(
                            {
                                "id": f"REVIEW-EXT-{name}-{address}",
                                "kind": "external_reference",
                                "status": "unresolved",
                                "source_cell": f"{name}!{address}",
                                "message": "外部工作簿引用尚未完成來源對應。",
                                "details": ref,
                            }
                        )
                if status == "error":
                    review_items.append(
                        {
                            "id": f"REVIEW-ERR-{name}-{address}",
                            "kind": "formula_error",
                            "status": "review_required",
                            "source_cell": f"{name}!{address}",
                            "message": "公式或其快取結果包含 Excel 錯誤狀態。",
                            "details": {
                                "formula": str(formula),
                                "cached_value": _json_value(cell.get("value")),
                            },
                        }
                    )
            elif status == "error":
                review_items.append(
                    {
                        "id": f"REVIEW-ERR-{name}-{address}",
                        "kind": "formula_error",
                        "status": "review_required",
                        "source_cell": f"{name}!{address}",
                        "message": "來源儲存格包含 Excel 錯誤狀態。",
                        "details": {"value": _json_value(cell.get("value"))},
                    }
                )

        model = {
            "schema_version": SCHEMA_VERSION,
            "profile": PROFILE,
            "sheet_meta": meta,
            "cells": cells,
            "visuals": legacy_sheet.get("visuals") or [],
            "sheet_risk_flags": sorted({item["kind"] for item in review_items}),
        }
        return model, formula_rows, edge_rows, review_items

    def _write_formula_catalog(self, rows: list[dict[str, Any]]) -> None:
        path = self.workspace / "normalized" / "formula_catalog.csv"
        fields = [
            "formula_id",
            "sheet",
            "cell",
            "source_cell",
            "formula_raw",
            "cached_value",
            "status",
            "function_names",
            "dependencies",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)

    def _write_dependency_edges(self, rows: list[dict[str, Any]]) -> None:
        path = self.workspace / "normalized" / "dependency_edges.csv"
        fields = [
            "source",
            "reference_kind",
            "target",
            "target_sheet",
            "target_range",
            "external_workbook",
            "status",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)

    def _write_review_items(self, items: list[dict[str, Any]]) -> None:
        path = self.workspace / "normalized" / "review_items.json"
        path.write_text(
            json.dumps(
                {
                    "schema_version": "fmeda-review-items-v1",
                    "items": items,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )

    def _write_summary(
        self,
        workbook: dict[str, Any],
        formulas: list[dict[str, Any]],
        edges: list[dict[str, Any]],
        reviews: list[dict[str, Any]],
    ) -> None:
        counts = Counter(item["kind"] for item in reviews)
        lines = [
            "# FMEDA Workbook Summary",
            "",
            "> 這是給審查者、主管與跨部門閱讀的摘要；公式完整資料位於 `formula_catalog.csv`，依賴位於 `dependency_edges.csv`。",
            "",
            f"**Schema**: `{workbook['schema_version']}`  ",
            f"**Profile**: `{workbook['profile']}`  ",
            f"**計算狀態**: `{workbook['source']['calculation_mode']}`  ",
            f"**工作表數**: {len(workbook['sheets'])}  ",
            f"**公式數**: {len(formulas)}  ",
            f"**依賴邊數**: {len(edges)}  ",
            f"**待審查項目**: {len(reviews)}",
            "",
            "## 審查重點",
            "",
        ]
        if counts:
            for kind, count in sorted(counts.items()):
                lines.append(f"- `{kind}`：{count} 件")
        else:
            lines.append("目前沒有偵測到需要人工審查的項目。")
        lines.extend(["", "## 工作表索引", "", "| # | 工作表 | 公式數 | 待審查 |", "|---:|---|---:|---:|"])
        for sheet in workbook["sheets"]:
            lines.append(
                f"| {sheet['index']} | `{sheet['name']}` | {sheet['formula_count']} | {sheet['review_count']} |"
            )
        lines.extend(
            [
                "",
                "## 閱讀規則",
                "",
                "公式結果若標示為 `cached_value`，代表讀取原始 Excel 保存的快取值，並不代表本工具重新計算過。`error`、`not_calculated`、空白與數值 0 必須分開解讀。",
                "",
                "外部工作簿引用若標示為 `unresolved`，代表保留原始引用字串但尚未完成來源對應；在人工確認前，不應自動改名或推導新的結果。",
                "",
            ]
        )
        (self.workspace / "normalized" / "Step03_summary.md").write_text(
            "\n".join(lines), encoding="utf-8"
        )

    @staticmethod
    def _build_import_report(manifest: dict[str, Any], reviews: list[dict[str, Any]]) -> str:
        lines = [
            "# FMEDA Import Report",
            "",
            f"**Source**: `{manifest['source_file']}`  ",
            f"**Source SHA-256**: `{manifest['source_sha256']}`  ",
            f"**Derived workbook**: `{manifest['derived_file']}`  ",
            f"**Schema**: `{manifest['schema_version']}`  ",
            f"**Profile**: `{manifest['profile']}`  ",
            "",
            "## Counts",
            "",
            f"- Sheets: {manifest['sheet_count']}",
            f"- Formulas: {manifest['formula_count']}",
            f"- Dependency edges: {manifest['dependency_edge_count']}",
            f"- Review items: {manifest['review_item_count']}",
            "",
            "## Safety notes",
            "",
            "原始檔只被讀取並複製到 source snapshot；任何可編輯操作應針對 derived workbook 或 Editor sidecar，不能覆蓋原始檔。",
            "",
            "本次輸出使用原始 Excel 的 cached values，尚未由本工具獨立重算。",
            "",
        ]
        if reviews:
            lines.extend(["## Review queue", ""])
            for item in reviews[:100]:
                lines.append(
                    f"- `{item['kind']}` `{item['source_cell']}` — {item['message']}（`{item['status']}`）"
                )
            if len(reviews) > 100:
                lines.append(f"- 其餘 {len(reviews) - 100} 項請查看 `normalized/review_items.json`。")
        return "\n".join(lines) + "\n"


# Public names: the explicit core name documents the new boundary, while the
# historical name remains available for callers from the earlier slice.
FmedaCoreWorkspaceBuilder = FmedaWorkspaceBuilder

__all__ = ["FmedaCoreWorkspaceBuilder", "FmedaWorkspaceBuilder"]
