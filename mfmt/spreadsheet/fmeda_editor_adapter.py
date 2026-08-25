"""Optional Markdown Editor adapter for the FMEDA text workspace.

The core workbook pipeline does not import this module.  The adapter consumes
core artifacts and writes Editor-specific Markdown and sidecar metadata while
preserving the same source revision and formula provenance.
"""

from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any

from .fmeda_workspace import (
    EDITOR_SIDECAR_SCHEMA,
    RELATIONS_SCHEMA,
    SCHEMA_VERSION,
    _safe_sheet_name,
)


class FmedaEditorAdapter:
    """Emit an optional Editor workspace from core FMEDA artifacts."""

    def emit(
        self,
        workspace: str | Path,
        workbook: dict[str, Any],
        sheets: list[dict[str, Any]],
        formulas: list[dict[str, Any]],
        reviews: list[dict[str, Any]],
        source_hash: str,
    ) -> dict[str, Any]:
        workspace_path = Path(workspace)
        editor_root = workspace_path / "editor"
        (editor_root / "sheets").mkdir(parents=True, exist_ok=True)

        formula_by_sheet: dict[str, list[dict[str, Any]]] = {}
        for row in formulas:
            formula_by_sheet.setdefault(str(row["sheet"]), []).append(row)
        review_by_sheet: dict[str, list[dict[str, Any]]] = {}
        for item in reviews:
            sheet = str(item["source_cell"]).split("!", 1)[0]
            review_by_sheet.setdefault(sheet, []).append(item)

        blocks: list[dict[str, Any]] = []
        groups: list[dict[str, Any]] = []
        documents: list[dict[str, Any]] = []
        order = 0
        for sheet in sheets:
            sheet_name = str(sheet["name"])
            stem = Path(
                str(sheet.get("md_file") or f"editor/sheets/{_safe_sheet_name(sheet_name, int(sheet['index']))}.md")
            ).name
            md_path = editor_root / "sheets" / stem
            sheet_formulas = formula_by_sheet.get(sheet_name, [])
            sheet_reviews = review_by_sheet.get(sheet_name, [])
            md_lines = [
                f"# {sheet_name}",
                "",
                "> FMEDA 工作表審查頁。此頁是 Editor 的人讀視圖；原始公式與完整依賴請回查 formula catalog。",
                "",
                f"**來源工作表**：`{sheet_name}`  ",
                f"**公式數**：{len(sheet_formulas)}  ",
                f"**待審查項目**：{len(sheet_reviews)}  ",
                f"**來源 SHA-256**：`{source_hash}`",
                "",
                "## 結果與公式索引",
                "",
                "| 儲存格 | 快取結果 | 狀態 | 公式 ID |",
                "|---|---:|---|---|",
            ]
            visible = sheet_formulas[:2_000]
            for row in visible:
                cached = str(row["cached_value"] or "").replace("|", "\\|")
                md_lines.append(
                    f"| `{row['source_cell']}` | {cached} | `{row['status']}` | `{row['formula_id']}` |"
                )
                md_lines.extend(
                    [
                        "",
                        f"### {row['formula_id']}",
                        "",
                        f"- 來源：`{row['source_cell']}`",
                        f"- 狀態：`{row['status']}`",
                        f"- 原始公式：`{row['formula_raw']}`",
                        f"- 函數：`{row['function_names'] or 'none'}`",
                        f"- 依賴：`{row['dependencies']}`",
                        "",
                    ]
                )
            if len(sheet_formulas) > 2_000:
                md_lines.extend(
                    [
                        "",
                        f"> 公式明細共 {len(sheet_formulas)} 筆；此頁先顯示前 2,000 筆，完整內容請查看 `../../normalized/formula_catalog.csv`。",
                        "",
                    ]
                )
            md_lines.extend(["## 待審查項目", ""])
            if sheet_reviews:
                for item in sheet_reviews:
                    md_lines.append(
                        f"- **{item['kind']}** `{item['source_cell']}`：{item['message']}（狀態：`{item['status']}`）"
                    )
            else:
                md_lines.append("目前沒有此工作表的待審查項目。")
            md_path.write_text("\n".join(md_lines) + "\n", encoding="utf-8")

            document_id = f"doc-{_safe_sheet_name(sheet_name, int(sheet['index']))}"
            documents.append(
                {
                    "id": document_id,
                    "relative_path": f"sheets/{stem}",
                    "name": stem,
                    "display_label": sheet_name,
                    "source_id": f"sheet:{sheet_name}",
                }
            )
            group_id = f"group-{document_id}"
            member_ids: list[str] = []
            summary_id = f"block-{document_id}-summary"
            member_ids.append(summary_id)
            blocks.append(
                {
                    "id": summary_id,
                    "ref_id": f"{sheet_name}:summary",
                    "source_id": f"sheet:{sheet_name}",
                    "reference_kind": "GRP",
                    "display_label": f"{sheet_name} summary",
                    "order": order,
                    "structure": "paragraph",
                    "semantic_kind": "note",
                    "markdown": f"# {sheet_name}\n\nFMEDA 工作表摘要與審查入口。",
                    "html": f"<h1>{html.escape(sheet_name)}</h1><p>FMEDA 工作表摘要與審查入口。</p>",
                    "editable_html": False,
                    "canvas_edit_mode": "markdown",
                    "ignored": False,
                    "group_id": group_id,
                    "properties": {
                        "source_revision": source_hash,
                        "editability": "read_only_source_summary",
                    },
                    "source_cell": f"{sheet_name}!__summary__",
                    "formula_id": None,
                    "editability": "read_only_source_summary",
                }
            )
            order += 1
            for row in visible:
                block_id = f"block-{row['formula_id']}"
                member_ids.append(block_id)
                blocks.append(
                    {
                        "id": block_id,
                        "ref_id": row["source_cell"],
                        "source_id": row["formula_id"],
                        "reference_kind": "BLK",
                        "display_label": row["source_cell"],
                        "order": order,
                        "structure": "paragraph",
                        "semantic_kind": "example",
                        "markdown": f"### {row['formula_id']}\n\n`{row['formula_raw']}`",
                        "html": f"<h3>{html.escape(row['formula_id'])}</h3><p><code>{html.escape(str(row['formula_raw']))}</code></p>",
                        "editable_html": False,
                        "canvas_edit_mode": "markdown",
                        "ignored": False,
                        "group_id": group_id,
                        "properties": {
                            "source_revision": source_hash,
                            "formula_status": str(row["status"]),
                            "cached_value": str(row["cached_value"]),
                        },
                        "source_cell": row["source_cell"],
                        "formula_id": row["formula_id"],
                        "editability": "read_only_formula",
                    }
                )
                order += 1
            groups.append(
                {
                    "id": group_id,
                    "ref_id": f"{sheet_name}:section",
                    "source_id": f"sheet:{sheet_name}",
                    "display_label": sheet_name,
                    "reference_kind": "GRP",
                    "group_type": "Section",
                    "block_ids": member_ids,
                    "properties": {
                        "section_title": sheet_name,
                        "section_level": "1",
                        "grouping_scope": "sheet",
                        "grouping_source": "fmeda_editor_adapter",
                    },
                }
            )

        sidecar = {
            "schema_version": EDITOR_SIDECAR_SCHEMA,
            "source_revision": {"sha256": source_hash, "schema_version": SCHEMA_VERSION},
            "documents": documents,
            "groups": groups,
            "blocks": blocks,
        }
        (editor_root / "blocks.sidecar.json").write_text(
            json.dumps(sidecar, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        (editor_root / "relations.json").write_text(
            json.dumps(
                {
                    "schema_version": RELATIONS_SCHEMA,
                    "source_revision": {"sha256": source_hash},
                    "relations": [],
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
        index_lines = [
            "# FMEDA Editor Workspace",
            "",
            "> 這是可選的 Editor adapter；核心純文字 workspace 不依賴此目錄。原始 Excel 位於 `../source/`，衍生 Excel 位於 `../derived/`。",
            "",
            f"**來源檔**：`{workbook['source']['filename']}`  ",
            f"**來源 SHA-256**：`{source_hash}`  ",
            f"**工作表數**：{len(sheets)}  ",
            f"**公式數**：{sum(int(sheet['formula_count']) for sheet in sheets)}",
            "",
            "## 工作表",
            "",
            "| # | 工作表 | Editor 文件 | 公式數 | 待審查 |",
            "|---:|---|---|---:|---:|",
        ]
        for sheet in sheets:
            name = str(sheet["name"]).replace("|", "\\|")
            md_name = Path(
                str(sheet.get("md_file") or f"editor/sheets/{_safe_sheet_name(str(sheet['name']), int(sheet['index']))}.md")
            ).name
            index_lines.append(
                f"| {sheet['index']} | {name} | [{md_name}](sheets/{md_name}) | {sheet['formula_count']} | {sheet['review_count']} |"
            )
        index_lines.extend(
            [
                "",
                "## 相關資料",
                "",
                "- [Step03_summary.md](../normalized/Step03_summary.md)",
                "- [formula_catalog.csv](../normalized/formula_catalog.csv)",
                "- [dependency_edges.csv](../normalized/dependency_edges.csv)",
                "- [review_items.json](../normalized/review_items.json)",
                "- [import-report.md](../reports/import-report.md)",
                "",
            ]
        )
        (editor_root / "index.md").write_text("\n".join(index_lines), encoding="utf-8")
        return {
            "enabled": True,
            "root": "editor",
            "sidecar": "editor/blocks.sidecar.json",
            "relations": "editor/relations.json",
            "schema_version": EDITOR_SIDECAR_SCHEMA,
        }


__all__ = ["FmedaEditorAdapter"]
