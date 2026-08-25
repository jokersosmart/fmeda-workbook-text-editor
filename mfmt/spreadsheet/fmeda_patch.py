"""Controlled patches for FMEDA derived workbooks.

This slice intentionally edits only a derived workbook copy.  Formula cells
are read-only; input changes require an explicit patch manifest and an
expected-value check.  The source snapshot is never opened for writing.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("", WORKBOOK_NS)
ET.register_namespace("r", REL_NS)

_CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")
_REV_RE = re.compile(r"\.rev-(\d+)(?=\.xlsx$)", re.IGNORECASE)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _tag(name: str) -> str:
    return f"{{{WORKBOOK_NS}}}{name}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _safe_stem(path: Path) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._") or "workbook"


def _normalize_zip_target(target: str) -> str:
    normalized = target.replace("\\", "/").lstrip("/")
    if not normalized.startswith("xl/"):
        normalized = f"xl/{normalized}"
    return normalized


def _child(parent: ET.Element, name: str) -> ET.Element | None:
    for child in parent:
        if _local(child.tag) == name:
            return child
    return None


def _cell_to_value(cell: ET.Element) -> Any:
    cell_type = cell.attrib.get("t")
    formula = _child(cell, "f")
    if formula is not None:
        formula_text = formula.text or ""
        return {"formula": f"={formula_text}"}
    if cell_type == "inlineStr":
        inline = _child(cell, "is")
        if inline is None:
            return ""
        return "".join(node.text or "" for node in inline.iter() if _local(node.tag) == "t")
    value_node = _child(cell, "v")
    raw = "" if value_node is None or value_node.text is None else value_node.text
    if cell_type == "b":
        return raw == "1"
    if cell_type == "e":
        return raw
    if raw == "":
        return None
    try:
        if "." in raw or "e" in raw.lower():
            return float(raw)
        return int(raw)
    except ValueError:
        return raw


def _same_value(left: Any, right: Any) -> bool:
    if left == right:
        return True
    if left is None or right is None:
        return False
    return str(left) == str(right)


def _value_to_nodes(value: Any) -> tuple[str | None, list[ET.Element]]:
    if value is None:
        return None, [ET.Element(_tag("v"))]
    if isinstance(value, bool):
        value_node = ET.Element(_tag("v"))
        value_node.text = "1" if value else "0"
        return "b", [value_node]
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        value_node = ET.Element(_tag("v"))
        value_node.text = repr(value)
        return None, [value_node]
    inline = ET.Element(_tag("is"))
    text_node = ET.SubElement(inline, _tag("t"))
    text_node.text = str(value)
    if str(value) != str(value).strip():
        text_node.set(f"{{{XML_SPACE}}}space", "preserve")
    return "inlineStr", [inline]


def _set_cell_value(cell: ET.Element, value: Any) -> None:
    for child in list(cell):
        if _local(child.tag) in {"f", "v", "is"}:
            cell.remove(child)
    cell_type, nodes = _value_to_nodes(value)
    if cell_type is None:
        cell.attrib.pop("t", None)
    else:
        cell.set("t", cell_type)
    for node in nodes:
        cell.append(node)


def _column_number(address: str) -> int:
    match = _CELL_RE.match(address.replace("$", ""))
    if not match:
        raise ValueError(f"invalid cell address: {address}")
    letters = match.group(1)
    number = 0
    for char in letters:
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def _worksheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets: dict[str, str] = {}
    for rel in relationships:
        if _local(rel.tag) != "Relationship":
            continue
        targets[rel.attrib.get("Id", "")] = _normalize_zip_target(rel.attrib.get("Target", ""))
    result: dict[str, str] = {}
    for sheet in workbook.iter(_tag("sheet")):
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{REL_NS}}}id", "")
        if name and rel_id in targets:
            result[name] = targets[rel_id]
    return result


def _find_cell(sheet_root: ET.Element, address: str) -> ET.Element | None:
    normalized = address.replace("$", "")
    for cell in sheet_root.iter(_tag("c")):
        if cell.attrib.get("r", "") == normalized:
            return cell
    return None


def _ensure_cell(sheet_root: ET.Element, address: str) -> ET.Element:
    existing = _find_cell(sheet_root, address)
    if existing is not None:
        return existing
    match = _CELL_RE.match(address.replace("$", ""))
    if not match:
        raise ValueError(f"invalid cell address: {address}")
    row_number = int(match.group(2))
    sheet_data = _child(sheet_root, "sheetData")
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, _tag("sheetData"))
    row = None
    for candidate in sheet_data:
        if _local(candidate.tag) == "row" and int(candidate.attrib.get("r", "0")) == row_number:
            row = candidate
            break
    if row is None:
        row = ET.Element(_tag("row"), {"r": str(row_number)})
        rows = [item for item in sheet_data if _local(item.tag) == "row"]
        insert_at = len(sheet_data)
        for index, candidate in enumerate(rows):
            if int(candidate.attrib.get("r", "0")) > row_number:
                insert_at = list(sheet_data).index(candidate)
                break
        sheet_data.insert(insert_at, row)
    cell = ET.Element(_tag("c"), {"r": match.group(1) + str(row_number)})
    target_column = _column_number(address)
    insert_at = len(row)
    for index, candidate in enumerate(row):
        if _local(candidate.tag) != "c":
            continue
        current_address = candidate.attrib.get("r", "")
        try:
            if _column_number(current_address) > target_column:
                insert_at = index
                break
        except ValueError:
            continue
    row.insert(insert_at, cell)
    return cell


def _replace_sheet_xml(
    source: Path,
    destination: Path,
    changes: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as archive:
        sheet_paths = _worksheet_paths(archive)
        grouped: dict[str, list[dict[str, Any]]] = {}
        for change in changes:
            sheet = str(change.get("sheet") or "").strip()
            if sheet not in sheet_paths:
                raise ValueError(f"sheet not found in workbook: {sheet}")
            grouped.setdefault(sheet, []).append(change)

        updated: list[dict[str, Any]] = []
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=destination.parent) as temp:
            temp_path = Path(temp.name)
        try:
            with zipfile.ZipFile(temp_path, "w") as output:
                for info in archive.infolist():
                    payload = archive.read(info.filename)
                    if info.filename in {sheet_paths[sheet] for sheet in grouped}:
                        sheet_name = next(name for name, path in sheet_paths.items() if path == info.filename)
                        root = ET.fromstring(payload)
                        for change in grouped[sheet_name]:
                            address = str(change.get("cell") or "").strip().upper()
                            if not _CELL_RE.match(address):
                                raise ValueError(f"invalid cell address: {address}")
                            cell = _ensure_cell(root, address)
                            current = _cell_to_value(cell)
                            if isinstance(current, dict) and "formula" in current:
                                raise ValueError(
                                    f"formula cell is read-only: {sheet_name}!{address}"
                                )
                            if "expected_old_value" in change and not _same_value(
                                current, change.get("expected_old_value")
                            ):
                                raise ValueError(
                                    f"patch conflict at {sheet_name}!{address}: "
                                    f"expected {change.get('expected_old_value')!r}, got {current!r}"
                                )
                            new_value = change.get("new_value")
                            _set_cell_value(cell, new_value)
                            updated.append(
                                {
                                    "sheet": sheet_name,
                                    "cell": address,
                                    "old_value": current,
                                    "new_value": new_value,
                                    "editability": "input",
                                    "status": "updated",
                                }
                            )
                        payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    if changes and info.filename == "xl/workbook.xml":
                        root = ET.fromstring(payload)
                        calc_pr = next(
                            (node for node in root if _local(node.tag) == "calcPr"),
                            None,
                        )
                        if calc_pr is None:
                            calc_pr = ET.SubElement(root, _tag("calcPr"))
                        calc_pr.set("calcMode", "auto")
                        calc_pr.set("fullCalcOnLoad", "1")
                        calc_pr.set("forceFullCalc", "1")
                        payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    output.writestr(info, payload)
            shutil.move(temp_path, destination)
        finally:
            if temp_path.exists():
                temp_path.unlink()
    return updated


def _next_revision_path(workspace: Path, base: Path) -> tuple[int, Path]:
    derived_dir = workspace / "derived"
    derived_dir.mkdir(parents=True, exist_ok=True)
    revisions: list[int] = []
    for candidate in derived_dir.glob("*.xlsx"):
        match = _REV_RE.search(candidate.name)
        if match:
            revisions.append(int(match.group(1)))
    revision = max(revisions or [0]) + 1
    stem = re.sub(r"\.rev-\d+$", "", base.stem, flags=re.IGNORECASE)
    return revision, derived_dir / f"{stem}.rev-{revision:03d}.xlsx"


class FmedaPatchApplier:
    """Apply explicit, formula-safe changes to a derived FMEDA workbook."""

    def __init__(self, workspace: str | Path):
        self.workspace = Path(workspace).expanduser().resolve()
        self.manifest_path = self.workspace / "manifest.json"
        if not self.manifest_path.is_file():
            raise FileNotFoundError(f"FMEDA workspace manifest not found: {self.manifest_path}")
        self.manifest = json.loads(self.manifest_path.read_text(encoding="utf-8"))
        self.source_hash = str(self.manifest.get("source_sha256") or "")
        if not self.source_hash:
            raise ValueError("workspace manifest is missing source_sha256")

    def apply(self, patch: dict[str, Any]) -> dict[str, Any]:
        if str(patch.get("schema_version") or "") != "fmeda-patch-v1":
            raise ValueError("patch schema_version must be fmeda-patch-v1")
        base_hash = str(patch.get("base_source_sha256") or "")
        if base_hash != self.source_hash:
            raise ValueError(
                "source revision conflict: patch base_source_sha256 does not match workspace"
            )
        changes = patch.get("changes") or []
        if not isinstance(changes, list):
            raise ValueError("patch changes must be a list")
        for change in changes:
            if not isinstance(change, dict):
                raise ValueError("each patch change must be an object")
            if not change.get("sheet") or not change.get("cell"):
                raise ValueError("each patch change requires sheet and cell")
            if change.get("editability") != "input":
                raise ValueError("each workbook change must declare editability=input")
            if "new_value" not in change:
                raise ValueError("each patch change requires new_value")

        derived_relative = str(self.manifest.get("derived_file") or "")
        base_derived = self.workspace / derived_relative
        if not base_derived.is_file():
            raise FileNotFoundError(f"derived workbook not found: {base_derived}")
        base_derived_hash = _sha256(base_derived)
        revision, destination = _next_revision_path(self.workspace, base_derived)
        updated = _replace_sheet_xml(base_derived, destination, changes)
        destination_hash = _sha256(destination)
        patch_id = str(patch.get("patch_id") or f"patch-rev-{revision:03d}")
        now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

        review_notes = patch.get("review_notes") or []
        if not isinstance(review_notes, list):
            raise ValueError("review_notes must be a list")
        review_path = self.workspace / "editor" / "review_notes.json"
        existing_notes: list[dict[str, Any]] = []
        if review_path.is_file():
            existing = json.loads(review_path.read_text(encoding="utf-8"))
            existing_notes = list(existing.get("notes") or [])
        normalized_notes = []
        for note in review_notes:
            if not isinstance(note, dict):
                raise ValueError("each review note must be an object")
            normalized_notes.append({**note, "patch_id": patch_id, "created_at": now})
        review_path.parent.mkdir(parents=True, exist_ok=True)
        review_path.write_text(
            json.dumps(
                {
                    "schema_version": "fmeda-review-notes-v1",
                    "source_sha256": self.source_hash,
                    "notes": existing_notes + normalized_notes,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )

        patch_record = {
            "schema_version": "fmeda-patch-manifest-v1",
            "patch_id": patch_id,
            "created_at": now,
            "base_source_sha256": self.source_hash,
            "base_derived_file": derived_relative,
            "base_derived_sha256": base_derived_hash,
            "derived_file": _relative(destination, self.workspace),
            "derived_sha256": destination_hash,
            "changes": updated,
            "review_note_count": len(normalized_notes),
            "status": "applied",
        }
        normalized_dir = self.workspace / "normalized"
        normalized_dir.mkdir(parents=True, exist_ok=True)
        (normalized_dir / f"patch_manifest.rev-{revision:03d}.json").write_text(
            json.dumps(patch_record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )

        report_lines = [
            f"# FMEDA Derived Export Report — rev-{revision:03d}",
            "",
            f"**Patch**: `{patch_id}`  ",
            f"**Source SHA-256**: `{self.source_hash}`  ",
            f"**Base derived**: `{derived_relative}`  ",
            f"**New derived**: `{patch_record['derived_file']}`  ",
            f"**New derived SHA-256**: `{destination_hash}`  ",
            "",
            "## Changes",
            "",
            "| Sheet | Cell | Old value | New value | Status |",
            "|---|---|---|---|---|",
        ]
        for change in updated:
            report_lines.append(
                f"| `{change['sheet']}` | `{change['cell']}` | "
                f"`{change['old_value']}` | `{change['new_value']}` | `{change['status']}` |"
            )
        report_lines.extend(
            [
                "",
                f"Review notes added: **{len(normalized_notes)}**",
                "",
                "> 公式儲存格在此 slice 中是唯讀的；本報告只描述通過 expected-value 與 source revision 檢查的 input patch。",
                "",
            ]
        )
        (self.workspace / "reports" / f"export-report.rev-{revision:03d}.md").write_text(
            "\n".join(report_lines), encoding="utf-8"
        )
        return patch_record


def _relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()
