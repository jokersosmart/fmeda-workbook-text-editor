from __future__ import annotations

import hashlib
import json
from pathlib import Path

from openpyxl import Workbook

from mfmt.fmeda_cli import main as fmeda_cli_main
from mfmt.spreadsheet.fmeda_workspace import FmedaCoreWorkspaceBuilder


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _make_fixture(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "FMEDA"
    sheet["A1"] = 10
    sheet["B1"] = "=A1*2"
    sheet["C1"] = "=Missing!A1"
    workbook.save(path)


def test_core_builder_is_editor_independent_by_default(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    workspace = tmp_path / "core-workspace"
    _make_fixture(source)
    source_hash = _sha256(source)

    manifest = FmedaCoreWorkspaceBuilder(source, workspace).build()

    assert manifest["editor"] is None
    assert not (workspace / "editor").exists()
    assert (workspace / "normalized" / "Step03_workbook.json").is_file()
    assert (workspace / "normalized" / "Step03_summary.md").is_file()
    assert (workspace / "normalized" / "formula_catalog.csv").is_file()
    assert (workspace / "normalized" / "dependency_edges.csv").is_file()
    assert (workspace / "reports" / "import-report.md").is_file()
    assert manifest["source_sha256"] == source_hash
    assert _sha256(source) == source_hash


def test_editor_adapter_is_opt_in_and_shares_core_provenance(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    workspace = tmp_path / "editor-workspace"
    _make_fixture(source)
    source_hash = _sha256(source)

    manifest = FmedaCoreWorkspaceBuilder(
        source, workspace, include_editor=True
    ).build()

    assert manifest["editor"]["enabled"] is True
    assert (workspace / "editor" / "index.md").is_file()
    assert (workspace / "editor" / "blocks.sidecar.json").is_file()
    sidecar = json.loads(
        (workspace / "editor" / "blocks.sidecar.json").read_text(encoding="utf-8")
    )
    assert sidecar["source_revision"]["sha256"] == source_hash
    assert manifest["source_sha256"] == source_hash
    assert (workspace / "normalized" / "Step03_workbook.json").is_file()


def test_cli_defaults_to_core_and_can_opt_in_to_editor(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.xlsx"
    _make_fixture(source)

    core_workspace = tmp_path / "cli-core"
    monkeypatch.setattr(
        "sys.argv",
        ["fmeda-workspace", str(source), "--output-dir", str(core_workspace)],
    )
    assert fmeda_cli_main() == 0
    assert not (core_workspace / "editor").exists()
    assert (core_workspace / "readable" / "index.md").is_file()

    editor_workspace = tmp_path / "cli-editor"
    monkeypatch.setattr(
        "sys.argv",
        [
            "fmeda-workspace",
            str(source),
            "--output-dir",
            str(editor_workspace),
            "--adapter",
            "editor",
        ],
    )
    assert fmeda_cli_main() == 0
    assert (editor_workspace / "editor" / "index.md").is_file()
    assert (editor_workspace / "readable" / "index.md").is_file()
