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
READABLE_SCHEMA = "fmeda-readable-v1"
MAX_READABLE_FORMULAS_PER_SHEET = 120
MAX_READABLE_INPUTS_PER_SHEET = 120

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
        readable_manifest = self._write_readable_workspace(
            workbook_model,
            sheet_models,
            formula_rows,
            edge_rows,
            review_items,
            source_hash,
        )
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
            "readable": readable_manifest,
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
            "**Readable 入口**: [readable/index.md](../readable/index.md)  ",
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

    def _write_readable_workspace(
        self,
        workbook: dict[str, Any],
        sheet_models: list[dict[str, Any]],
        formulas: list[dict[str, Any]],
        edges: list[dict[str, Any]],
        reviews: list[dict[str, Any]],
        source_hash: str,
    ) -> dict[str, Any]:
        """Write the human-first core reading layer without duplicating truth.

        The machine-readable normalized artifacts remain authoritative. This
        layer adds navigation, bounded samples, and explanations so a reader
        does not need to open a 500k-formula CSV before understanding the
        workbook status.
        """
        root = self.workspace / "readable"
        sheets_root = root / "sheets"
        sheets_root.mkdir(parents=True, exist_ok=True)
        formulas_by_sheet: dict[str, list[dict[str, Any]]] = {}
        for row in formulas:
            formulas_by_sheet.setdefault(str(row["sheet"]), []).append(row)
        reviews_by_sheet: dict[str, list[dict[str, Any]]] = {}
        for item in reviews:
            source_cell = str(item.get("source_cell", ""))
            sheet_name = source_cell.split("!", 1)[0] if "!" in source_cell else ""
            reviews_by_sheet.setdefault(sheet_name, []).append(item)
        edges_by_sheet: dict[str, list[dict[str, Any]]] = {}
        for edge in edges:
            source_cell = str(edge.get("source", ""))
            sheet_name = source_cell.split("!", 1)[0] if "!" in source_cell else ""
            edges_by_sheet.setdefault(sheet_name, []).append(edge)

        readable_sheets: list[dict[str, Any]] = []
        for sheet in sheet_models:
            meta = dict(sheet.get("sheet_meta") or {})
            name = str(meta.get("name") or "Sheet")
            index = int(meta.get("index") or len(readable_sheets) + 1)
            safe_name = _safe_sheet_name(name, index)
            output_name = f"{safe_name}.md"
            output_path = sheets_root / output_name
            cells = list((sheet.get("cells") or {}).values())
            sheet_formulas = formulas_by_sheet.get(name, [])
            sheet_reviews = reviews_by_sheet.get(name, [])
            sheet_edges = edges_by_sheet.get(name, [])
            status_counts = Counter(
                str(cell.get("calculation_status") or "unknown") for cell in cells
            )
            inputs = [
                cell
                for cell in cells
                if not cell.get("formula")
                and cell.get("value") is not None
                and not _is_error(cell.get("value"))
            ]
            formula_sample = sheet_formulas[:MAX_READABLE_FORMULAS_PER_SHEET]
            input_sample = sorted(inputs, key=lambda cell: str(cell.get("address", "")))[:MAX_READABLE_INPUTS_PER_SHEET]
            function_counts = Counter(
                function_name
                for row in sheet_formulas
                for function_name in str(row.get("function_names") or "").split(";")
                if function_name
            )
            external_dependency_count = sum(
                1 for edge in sheet_edges if edge.get("reference_kind") == "external_reference"
            )
            formula_error_count = sum(
                1 for item in sheet_reviews if item.get("kind") == "formula_error"
            )
            merged = meta.get("merged_ranges") or meta.get("merged_cells") or []
            dimension = meta.get("dimension") or meta.get("dimensions") or "not recorded"
            if isinstance(dimension, dict):
                max_row = dimension.get("max_row")
                max_col = dimension.get("max_col")
                if max_row is not None and max_col is not None:
                    dimension_label = f"rows 1–{max_row}, columns 1–{max_col}"
                else:
                    dimension_label = json.dumps(dimension, ensure_ascii=False, sort_keys=True)
            else:
                dimension_label = str(dimension)
            sheet_json = next(
                (
                    entry.get("json_file")
                    for entry in workbook.get("sheets", [])
                    if str(entry.get("name")) == name
                ),
                "",
            )

            def text(value: Any) -> str:
                return str(_json_value(value) if value is not None else "—").replace("|", "\\|").replace("\n", " ")

            lines = [
                f"# {name}",
                "",
                "> 這一頁先說明工作表狀態，再列出關鍵輸入與公式樣本；完整資料請沿著 provenance 連結回查。",
                "",
                "## 一、先看結論",
                "",
                f"- 工作表位置：第 {index} 張",
                f"- 工作表尺寸：`{dimension_label}`",
                f"- 合併儲存格範圍：{len(merged) if isinstance(merged, (list, tuple)) else text(merged)}",
                f"- 公式數：{len(sheet_formulas)}",
                f"- 非空輸入／常數數：{len(inputs)}",
                f"- 待審查項目：{len(sheet_reviews)}",
                f"- 外部引用數：{external_dependency_count}",
                f"- 公式錯誤訊號數：{formula_error_count}",
                f"- 使用的公式函數：{', '.join(f'{name} ({count})' for name, count in sorted(function_counts.items())) or '未辨識到函數名稱'}",
                "",
                "### 狀態分布",
                "",
                "| 狀態 | 數量 | 怎麼解讀 |",
                "|---|---:|---|",
            ]
            status_meanings = {
                "input_or_constant": "可能是輸入或固定值，需依欄位語意判斷",
                "cached_value": "公式存在，值是來源 Excel 的快取結果",
                "error": "公式或來源包含錯誤狀態",
                "not_calculated": "公式存在但沒有快取結果",
                "empty": "空白儲存格",
                "unknown": "轉換器未提供狀態",
            }
            for status, count in sorted(status_counts.items()):
                lines.append(f"| `{status}` | {count} | {status_meanings.get(status, '請回查 workbook JSON') } |")
            lines.extend(["", "### 目前需要注意", ""])
            if sheet_reviews:
                for item in sheet_reviews[:20]:
                    lines.append(
                        f"- `{item.get('kind')}` `{item.get('source_cell')}`：{item.get('message')}（`{item.get('status')}`）"
                    )
                if len(sheet_reviews) > 20:
                    lines.append(
                        f"- 其餘 {len(sheet_reviews) - 20} 項請查看 `../review-queue.md` 或完整 `../../normalized/review_items.json`。"
                    )
            else:
                lines.append("目前沒有被核心 parser 標記的待審查項目。")

            lines.extend(
                [
                    "",
                    "## 二、關鍵輸入／常數（前 120 筆）",
                    "",
                    "| 儲存格 | 值 | 原始型別 | 來源 |",
                    "|---|---|---|---|",
                ]
            )
            for cell in input_sample:
                address = str(cell.get("address") or "")
                lines.append(
                    f"| `{name}!{address}` | `{text(cell.get('value'))}` | `{text(cell.get('data_type'))}` | `source: {source_hash}` |"
                )
            if len(inputs) > len(input_sample):
                full_sheet_path = f"../../{sheet_json}" if sheet_json else "../../normalized/sheets/<sheet>.json"
                lines.extend(
                    [
                        "",
                        f"> 這一頁顯示前 {len(input_sample)} 筆；完整儲存格請查看 [工作表 JSON]({full_sheet_path})。",
                        "",
                    ]
                )
            elif not input_sample:
                lines.append("| — | （沒有非空輸入／常數） | — | — |")

            lines.extend(
                [
                    "",
                    "## 三、公式計算摘要（前 120 筆）",
                    "",
                    "> 欄位對應：`formula_raw` 是原始公式；`cached_value` 是來源快取結果；`calculation_status` 是公式狀態；`formula_id` 是可回查索引。",
                    "",
                    "| 儲存格 | 公式 | 快取結果 | 狀態 | Formula ID |",
                    "|---|---|---|---|---|",
                ]
            )
            for row in formula_sample:
                lines.append(
                    f"| `{row['source_cell']}` | `{text(row.get('formula_raw'))}` | `{text(row.get('cached_value'))}` | `{row.get('status')}` | `{row.get('formula_id')}` |"
                )
            if len(sheet_formulas) > len(formula_sample):
                lines.extend(
                    [
                        "",
                        f"> 公式數量為 {len(sheet_formulas)}，此頁只顯示前 {len(formula_sample)} 筆，避免大型工作表變成無法閱讀的長頁面。完整公式請查閱 `../../normalized/formula_catalog.csv`。",
                        "",
                    ]
                )
            elif not formula_sample:
                lines.append("| — | （沒有公式） | — | — | — |")

            lines.extend(
                [
                    "",
                    "## 四、如何追溯",
                    "",
                    f"- 完整工作表 JSON：[normalized/sheets/{Path(sheet_json).name if sheet_json else '<sheet>.json'}](../../{sheet_json or 'normalized/sheets/<sheet>.json'})",
                    "- 完整公式目錄：[normalized/formula_catalog.csv](../../normalized/formula_catalog.csv)",
                    "- 依賴索引：[normalized/dependency_edges.csv](../../normalized/dependency_edges.csv)",
                    "- 審查佇列：[readable/review-queue.md](../review-queue.md)",
                    f"- `source_sha256`：`{source_hash}`",
                    "",
                ]
            )
            output_path.write_text("\n".join(lines), encoding="utf-8")
            readable_sheets.append(
                {
                    "index": index,
                    "name": name,
                    "file": f"readable/sheets/{output_name}",
                    "formula_count": len(sheet_formulas),
                    "input_count": len(inputs),
                    "review_count": len(sheet_reviews),
                    "external_dependency_count": external_dependency_count,
                    "formula_error_count": formula_error_count,
                    "formula_detail_limit": MAX_READABLE_FORMULAS_PER_SHEET,
                }
            )

        index_lines = [
            "# FMEDA Core Reading Index",
            "",
            "> 先看這裡：這是給審查者、主管與跨部門人員的入口。先看整體狀態，再按工作表深入；完整公式與依賴永遠回查 normalized 目錄。",
            "",
            "## 0. 這份 workspace 是什麼",
            "",
            "- 這是核心純文字閱讀層，不依賴 Markdown Editor。",
            "- 原始 Excel 只讀取並保存於 `../source/`；衍生工作版本位於 `../derived/`。",
            f"- 來源檔：`{workbook['source']['filename']}`",
            f"- `source_sha256`：`{source_hash}`",
            f"- 計算狀態：`{workbook['source']['calculation_mode']}`",
            "",
            "## 1. 先看總體數字",
            "",
            "| 指標 | 數值 | 回查位置 |",
            "|---|---:|---|",
            f"| 工作表 | {len(workbook.get('sheets', []))} | `manifest.json` |",
            f"| 公式 | {len(formulas)} | `../normalized/formula_catalog.csv` |",
            f"| 依賴邊 | {len(edges)} | `../normalized/dependency_edges.csv` |",
            f"| 待審查項目 | {len(reviews)} | [`review-queue.md`](review-queue.md) |",
            "",
            "## 2. 審查順序",
            "",
            "1. 先看 [`review-queue.md`](review-queue.md)，確認錯誤、外部引用與未計算項目。",
            "2. 再看 [`formula-guide.md`](formula-guide.md)，了解 `cached_value`、`error` 與 `not_calculated` 的差異。",
            "3. 依工作表閱讀下表；需要公式細節時，再回查 `formula_catalog.csv`。",
            "4. 需要更深層的依賴追蹤時，使用 `dependency_edges.csv` 與工作表 JSON。",
            "",
            "## 3. 工作表索引",
            "",
            "| # | 工作表 | 閱讀頁 | 公式 | 輸入／常數 | 待審查 |",
            "|---:|---|---|---:|---:|---:|",
        ]
        for sheet in readable_sheets:
            index_lines.append(
                f"| {sheet['index']} | `{sheet['name']}` | [{Path(sheet['file']).name}](sheets/{Path(sheet['file']).name}) | {sheet['formula_count']} | {sheet['input_count']} | {sheet['review_count']} |"
            )
        priority_sheets = sorted(
            readable_sheets,
            key=lambda sheet: (
                -sheet["review_count"],
                -sheet["formula_error_count"],
                -sheet["external_dependency_count"],
                sheet["index"],
            ),
        )
        priority_lines = [
            "## 1.5. 優先注意工作表",
            "",
            "> 以下只依核心 parser 的待審查、公式錯誤與外部依賴訊號排序；它不是 FMEDA 安全關鍵性或失效率的語意判讀。",
            "",
            "| 順位 | 工作表 | 待審查 | 公式錯誤訊號 | 外部依賴 |",
            "|---:|---|---:|---:|---:|",
        ]
        for rank, sheet in enumerate(priority_sheets[:10], start=1):
            priority_lines.append(
                f"| {rank} | [{sheet['name']}](sheets/{Path(sheet['file']).name}) | {sheet['review_count']} | {sheet['formula_error_count']} | {sheet['external_dependency_count']} |"
            )
        priority_lines.append("")
        insert_at = index_lines.index("## 2. 審查順序")
        index_lines[insert_at:insert_at] = priority_lines
        index_lines.extend(
            [
                "",
                "## 4. 核心資料連結",
                "",
                "- [workbook-v2 JSON](../normalized/Step03_workbook.json)",
                "- [formula catalog](../normalized/formula_catalog.csv)",
                "- [dependency edges](../normalized/dependency_edges.csv)",
                "- [review items](../normalized/review_items.json)",
                "- [import report](../reports/import-report.md)",
                "",
            ]
        )
        (root / "index.md").write_text("\n".join(index_lines), encoding="utf-8")

        review_counts = Counter(str(item.get("kind", "unknown")) for item in reviews)
        review_status_counts = Counter(str(item.get("status", "unknown")) for item in reviews)
        review_lines = [
            "# Review Queue",
            "",
            "> 這份佇列只列出需要人判斷或回查的事項；它不是把錯誤清掉，而是把不確定性集中到可操作的位置。",
            "",
            f"- `source_sha256`：`{source_hash}`",
            f"- 總項目：{len(reviews)}",
            "",
            "## 按類型統計",
            "",
            "| 類型 | 數量 |",
            "|---|---:|",
        ]
        for kind, count in sorted(review_counts.items()):
            review_lines.append(f"| `{kind}` | {count} |")
        review_lines.extend(["", "## 按狀態統計", "", "| 狀態 | 數量 |", "|---|---:|"])
        for status, count in sorted(review_status_counts.items()):
            review_lines.append(f"| `{status}` | {count} |")
        review_lines.extend(["", "## 先看這些項目（最多 200 筆）", ""])
        if reviews:
            for item in reviews[:200]:
                review_lines.append(
                    f"- `{item.get('kind')}` `{item.get('source_cell')}`：{item.get('message')}（狀態：`{item.get('status')}`）"
                )
            if len(reviews) > 200:
                review_lines.append(f"- 其餘 {len(reviews) - 200} 項請查看 `../normalized/review_items.json`。")
        else:
            review_lines.append("目前沒有待審查項目。")
        (root / "review-queue.md").write_text("\n".join(review_lines) + "\n", encoding="utf-8")

        formula_guide = "\n".join(
            [
                "# Formula and Result Guide",
                "",
                "> 這一頁說明如何閱讀公式與結果；它不會把 cached value 冒充成最新重算值。",
                "",
                "## 欄位意義",
                "",
                "| 欄位 | 意義 |",
                "|---|---|",
                "| `formula_raw` | 原始 Excel 公式文字，作為最重要的公式證據。 |",
                "| `cached_value` | Excel 檔案保存的公式快取值，不代表本工具重新計算。 |",
                "| `calculation_status` | `cached_value`、`error`、`not_calculated` 等狀態。 |",
                "| `formula_id` | 讓 Markdown、JSON、CSV 與 Editor adapter 共用的穩定索引。 |",
                "| `dependency_edges` | 公式引用的同表、跨表或外部工作簿範圍。 |",
                "| `source_sha256` | 對應到來源檔版本，避免不同檔案的結果被混用。 |",
                "",
                "## 狀態判讀",
                "",
                "- `cached_value`：公式存在且有來源快取結果；如果需要新結果，必須另行執行受控重算。",
                "- `error`：公式或來源結果是 Excel 錯誤，不能當成 0 或空白。",
                "- `not_calculated`：公式存在，但檔案沒有保存計算結果。",
                "- `input_or_constant`：非公式內容；是否可編輯要由 profile 或 patch contract 決定。",
                "- `unresolved`：外部引用尚未完成來源對應，不能自行猜測。",
                "",
                "## 建議閱讀方法",
                "",
                "1. 先看工作表頁面的摘要與 review queue。",
                "2. 再從 `formula_id` 回查完整 formula catalog。",
                "3. 再從 `dependency_edges.csv` 確認公式是否依賴外部檔案。",
                "4. 最後才判斷是否需要 Calc 重算或 FMEDA 工程師審查。",
                "",
            ]
        )
        (root / "formula-guide.md").write_text(formula_guide, encoding="utf-8")

        readable_manifest = {
            "schema_version": READABLE_SCHEMA,
            "source_sha256": source_hash,
            "source_file": workbook["source"]["source_file"],
            "calculation_mode": workbook["source"]["calculation_mode"],
            "index": "readable/index.md",
            "review_queue": "readable/review-queue.md",
            "formula_guide": "readable/formula-guide.md",
            "sheet_count": len(readable_sheets),
            "formula_count": len(formulas),
            "dependency_edge_count": len(edges),
            "review_item_count": len(reviews),
            "formula_detail_limit": MAX_READABLE_FORMULAS_PER_SHEET,
            "input_detail_limit": MAX_READABLE_INPUTS_PER_SHEET,
            "sheets": readable_sheets,
        }
        (root / "manifest.json").write_text(
            json.dumps(readable_manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return {
            "schema_version": READABLE_SCHEMA,
            "index": "readable/index.md",
            "manifest": "readable/manifest.json",
            "review_queue": "readable/review-queue.md",
            "formula_guide": "readable/formula-guide.md",
            "sheet_count": len(readable_sheets),
            "formula_count": len(formulas),
            "review_item_count": len(reviews),
        }

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
            "**Readable 入口**: [readable/index.md](../readable/index.md)  ",
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
