"""Excel-compatible recalculation through a disposable LibreOffice profile."""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any

from .fmeda_external import (
    ExternalLinkResolutionError,
    bind_external_links,
    materialize_external_workbooks,
    resolve_external_links,
)
from xml.etree import ElementTree as ET

try:
    from lxml import etree as LET
except ImportError:  # pragma: no cover - exercised only in minimal environments
    LET = None

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _iter_xml(source: Any):
    if LET is not None:
        return LET.iterparse(
            source,
            events=("end",),
            tag=(f"{{{MAIN_NS}}}c",),
            huge_tree=True,
        )
    return ET.iterparse(source, events=("end",))


def _count_formula_results(path: Path) -> tuple[int, int]:
    formula_count = 0
    cached_result_count = 0
    with zipfile.ZipFile(path, "r") as archive:
        worksheet_paths = sorted(
            name
            for name in archive.namelist()
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        )
        for worksheet_path in worksheet_paths:
            with archive.open(worksheet_path, "r") as stream:
                for _, cell in _iter_xml(stream):
                    has_formula = any(
                        child.tag.rsplit("}", 1)[-1] == "f" for child in cell
                    )
                    if not has_formula:
                        cell.clear()
                        continue
                    formula_count += 1
                    value_node = next(
                        (child for child in cell if child.tag.rsplit("}", 1)[-1] == "v"),
                        None,
                    )
                    if value_node is not None and value_node.text is not None:
                        cached_result_count += 1
                    cell.clear()
    return formula_count, cached_result_count


class LibreOfficeRecalculator:
    """Recalculate a workbook without ever writing to the source path."""

    def __init__(
        self,
        source: str | Path,
        output: str | Path,
        *,
        soffice: str = "soffice",
        timeout_seconds: int = 900,
        external_workbooks: list[str | Path] | None = None,
        external_mode: str = "bind",
        external_source_kind: str = "unclassified",
    ):
        self.source = Path(source).expanduser().resolve()
        self.output = Path(output).expanduser().resolve()
        self.soffice = soffice
        self.timeout_seconds = timeout_seconds
        self.external_workbooks = [
            Path(value).expanduser().resolve() for value in (external_workbooks or [])
        ]
        if external_mode not in {"bind", "materialize"}:
            raise ValueError("external_mode must be 'bind' or 'materialize'")
        if external_source_kind not in {"unclassified", "synthetic-fixture", "production"}:
            raise ValueError(
                "external_source_kind must be 'unclassified', 'synthetic-fixture', or 'production'"
            )
        self.external_mode = external_mode
        self.external_source_kind = external_source_kind
        if not self.source.is_file():
            raise FileNotFoundError(f"source workbook not found: {self.source}")
        if self.source.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("LibreOfficeRecalculator currently supports .xlsx and .xlsm")

    def run(self) -> dict[str, Any]:
        started = time.monotonic()
        source_hash_before = _sha256(self.source)
        self.output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="fmeda-recalc-") as temp_dir:
            temp_root = Path(temp_dir)
            profile = temp_root / "profile"
            input_copy = temp_root / self.source.name
            converted_dir = temp_root / "converted"
            converted_dir.mkdir()
            shutil.copy2(self.source, input_copy)
            external_report: dict[str, Any] = {"status": "NOT_REQUESTED", "links": []}
            if self.external_workbooks:
                try:
                    if self.external_mode == "bind":
                        external_report = bind_external_links(
                            input_copy,
                            input_copy.with_name(f"{input_copy.stem}.bound{input_copy.suffix}"),
                            self.external_workbooks,
                        )
                        bound_copy = input_copy.with_name(f"{input_copy.stem}.bound{input_copy.suffix}")
                    else:
                        external_report = materialize_external_workbooks(
                            input_copy,
                            input_copy.with_name(f"{input_copy.stem}.materialized{input_copy.suffix}"),
                            self.external_workbooks,
                        )
                        bound_copy = input_copy.with_name(
                            f"{input_copy.stem}.materialized{input_copy.suffix}"
                        )
                    input_copy = bound_copy
                    external_report["source_kind"] = self.external_source_kind
                except ExternalLinkResolutionError:
                    raise
            elif self.source.suffix.lower() in {".xlsx", ".xlsm"}:
                external_report = {
                    "status": "UNRESOLVED_NOT_SUPPLIED",
                    "source_kind": self.external_source_kind,
                    "links": [
                        {
                            "status": item.status,
                            "index": item.index,
                            "original_target": item.original_target,
                            "sheet_names": list(item.sheet_names),
                            "reason": item.reason,
                        }
                        for item in resolve_external_links(self.source, [])
                    ],
                }
            command = [
                self.soffice,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile.as_uri()}",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(converted_dir),
                str(input_copy),
            ]
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds,
                check=False,
            )
            if completed.returncode != 0:
                raise RuntimeError(
                    "LibreOffice recalculation failed "
                    f"(exit={completed.returncode}): {completed.stderr.strip() or completed.stdout.strip()}"
                )
            converted = converted_dir / self.source.name
            if not converted.is_file():
                candidates = list(converted_dir.glob("*.xlsx"))
                if not candidates:
                    raise RuntimeError(
                        "LibreOffice completed without producing an xlsx output: "
                        f"{completed.stdout.strip()}"
                    )
                converted = candidates[0]
            shutil.copy2(converted, self.output)

        source_hash_after = _sha256(self.source)
        formula_count, cached_result_count = _count_formula_results(self.output)
        return {
            "schema_version": "fmeda-recalculation-report-v1",
            "status": "RECALCULATED",
            "engine": "libreoffice-calc",
            "source_file": self.source.name,
            "output_file": self.output.name,
            "source_sha256_before": source_hash_before,
            "source_sha256_after": source_hash_after,
            "output_sha256": _sha256(self.output),
            "formula_count": formula_count,
            "cached_result_count": cached_result_count,
            "external_resolution": external_report,
            "duration_seconds": round(time.monotonic() - started, 3),
        }
