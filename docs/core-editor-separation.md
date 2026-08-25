# Core / Editor Separation

## Purpose

The FMEDA text workspace has two layers with different responsibilities. The core layer is the source-safe Excel-to-text pipeline. The Editor layer is an optional adapter for collaborative reading, review notes, and controlled editing.

## Core-only workflow

```text
source.xlsx
    ↓
FmedaCoreWorkspaceBuilder
    ├── source snapshot + SHA-256
    ├── derived rev-001.xlsx
    ├── workbook-v2 JSON
    ├── formula_catalog.csv
    ├── dependency_edges.csv
    ├── review_items.json
    ├── Step03_summary.md
    ├── readable/index.md
    ├── readable/sheets/*.md
    ├── readable/review-queue.md
    ├── readable/formula-guide.md
    ├── readable/manifest.json
    └── reports/import-report.md
```

The core workflow does not create an `editor/` directory and does not import the Editor adapter during normal CLI startup. It does create the independent `readable/` layer, whose index-first Markdown is intended for reviewers, managers, and cross-functional readers. It can be used in scripts, CI jobs, batch processing, or environments where the Editor is not installed.

```powershell
fmeda-workspace input.xlsx --output-dir out/core-workspace
```

## Optional Editor adapter workflow

```text
core workbook-v2 + provenance
    ↓
readable/ (human-first core view)
    ↓
FmedaEditorAdapter
    ├── editor/index.md
    ├── editor/sheets/*.md
    ├── editor/blocks.sidecar.json
    └── editor/relations.json
```

The adapter is explicitly enabled:

```powershell
fmeda-workspace input.xlsx \
  --output-dir out/editor-workspace \
  --adapter editor
```

The adapter reads the core workbook model and writes Editor-specific files. The readable layer and the adapter both point back to the same normalized artifacts; neither creates a second formula truth. `source_sha256`, `source_cell`, `formula_id`, and source revision metadata remain aligned with the core artifacts.

## Python API

```python
from mfmt.spreadsheet.fmeda_workspace import FmedaCoreWorkspaceBuilder

manifest = FmedaCoreWorkspaceBuilder(source, output).build()
```

The historical `FmedaWorkspaceBuilder` name remains available as a compatibility alias. Its default is now core-only. Existing code that explicitly needs the Editor workspace can use `include_editor=True`:

```python
from mfmt.spreadsheet.fmeda_workspace import FmedaWorkspaceBuilder

manifest = FmedaWorkspaceBuilder(source, output, include_editor=True).build()
```

## Contract

The core is authoritative for source evidence, formulas, cached values, errors, dependencies, validation, patches, and derived revisions. `readable/` is a bounded human view: it shows summary, parser-signal priority, and samples, while complete details remain in `normalized/`. The adapter is authoritative only for Editor presentation and collaboration metadata. If the adapter is absent, the core remains complete for import, reading, validation, and derived-output workflows.
