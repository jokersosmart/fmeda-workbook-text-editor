from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from openpyxl import Workbook

from mfmt.spreadsheet.fmeda_diff import FmedaRevisionValidator


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_workbook(path: Path, input_value: float, formula: str = "=A1*2", label: str = "OK") -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "FMEDA"
    sheet["A1"] = input_value
    sheet["B1"] = formula
    sheet["C1"] = label
    workbook.save(path)


def test_allowed_input_change_passes_and_is_rendered(tmp_path: Path) -> None:
    base = tmp_path / "rev-001.xlsx"
    target = tmp_path / "rev-002.xlsx"
    _make_workbook(base, 1.0)
    _make_workbook(target, 2.0)
    patch = {
        "schema_version": "fmeda-patch-manifest-v1",
        "base_derived_sha256": _sha256(base),
        "derived_sha256": _sha256(target),
        "changes": [
            {
                "sheet": "FMEDA",
                "cell": "A1",
                "old_value": 1,
                "new_value": 2,
                "editability": "input",
            }
        ],
    }

    validator = FmedaRevisionValidator(base, target, patch)
    report = validator.validate()

    assert report["status"] == "PASS"
    assert report["allowed_input_changes"] == 1
    assert report["blocking_change_count"] == 0
    rendered = validator.render_markdown(report)
    assert "allowed_input" in rendered
    assert "PASS" in rendered


def test_formula_change_is_blocking(tmp_path: Path) -> None:
    base = tmp_path / "rev-001.xlsx"
    target = tmp_path / "rev-002.xlsx"
    _make_workbook(base, 1.0, formula="=A1*2")
    _make_workbook(target, 1.0, formula="=A1*3")

    report = FmedaRevisionValidator(base, target).validate()

    assert report["status"] == "FAIL"
    assert any(item["kind"] == "formula_raw" and item["blocking"] for item in report["differences"])


def test_unexpected_value_is_blocking_even_when_another_input_is_allowed(tmp_path: Path) -> None:
    base = tmp_path / "rev-001.xlsx"
    target = tmp_path / "rev-002.xlsx"
    _make_workbook(base, 1.0, label="OK")
    _make_workbook(target, 2.0, label="CHANGED")
    patch = {
        "schema_version": "fmeda-patch-manifest-v1",
        "base_derived_sha256": _sha256(base),
        "derived_sha256": _sha256(target),
        "changes": [
            {
                "sheet": "FMEDA",
                "cell": "A1",
                "old_value": 1,
                "new_value": 2,
                "editability": "input",
            }
        ],
    }

    report = FmedaRevisionValidator(base, target, patch).validate()

    assert report["status"] == "FAIL"
    assert report["allowed_input_changes"] == 1
    assert any(item["kind"] == "unexpected_value" for item in report["differences"])


def test_provenance_hash_mismatch_is_blocking(tmp_path: Path) -> None:
    base = tmp_path / "rev-001.xlsx"
    target = tmp_path / "rev-002.xlsx"
    _make_workbook(base, 1.0)
    _make_workbook(target, 2.0)
    patch = {
        "schema_version": "fmeda-patch-manifest-v1",
        "base_derived_sha256": "wrong-base-hash",
        "derived_sha256": _sha256(target),
        "changes": [
            {
                "sheet": "FMEDA",
                "cell": "A1",
                "old_value": 1,
                "new_value": 2,
                "editability": "input",
            }
        ],
    }

    report = FmedaRevisionValidator(base, target, patch).validate()

    assert report["status"] == "FAIL"
    assert report["provenance_issues"][0]["kind"] == "base_hash_mismatch"
