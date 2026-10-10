# Phase 1 Excel Decomposition Implementation Plan

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** Build a Phase 1 framework that decomposes any supported Excel workbook into complete, stored, reviewable modules without prematurely deciding the final actuarial software structure.

**Architecture:** Use a pluggable state-machine pipeline: ingest workbook → inventory every workbook object → build dependency/source graph → classify Excel-native modules → collect human confirmations → persist artifacts → generate reports. Python generation and actuarial-domain refactoring are deliberately deferred until Phase 1 artifacts are trustworthy.

**Tech Stack:** Python 3.11+, openpyxl as the default workbook reader, Python stdlib OOXML package scanning for unsupported parts, optional formulas/xlcalculator behind plugin boundaries, Pydantic for artifact contracts, NetworkX for graph representation, Typer for CLI, Jinja2 for Markdown reports, XlsxWriter later for optional Excel reports.

---

## Product decision

Do **not** start with direct `Excel → Python` conversion. The first durable product asset should be a lossless-ish, auditable workbook decomposition layer.

The current uncertainty is: **what exactly is the final actuarial structure?** For now, Phase 1 should avoid over-answering that question. Instead, it should split Excel according to its original workbook structure and preserve enough information to later support multiple targets:

- faithful Python reference implementation;
- actuarial-domain Python model;
- API/batch runtime;
- validation/reconciliation package;
- governance/audit package;
- future Java/C#/Rust/SQL/WASM generators.

Phase 1 only needs light domain hints. It does not need to decide final `Product`, `Policy`, `AssumptionSet`, `ProjectionEngine`, or `Cashflow` abstractions yet.

## Core principle: cover all Excel content

The module split must cover the whole workbook, not only formulas that look actuarial. Phase 1 artifacts must preserve the original workbook hierarchy:

```text
workbook → sheet order → sheet → row/column → cell/range/object → source_location
```

Classification must never replace the raw inventory. It is an overlay on top of source-preserving workbook facts.

Every workbook artifact should land in one of these Phase 1 buckets:

1. **Workbook metadata** — file hash, format, author fields where available, calculation mode, workbook protection.
2. **Sheet inventory** — visible/hidden/very hidden sheets, dimensions, tab colors, protection state.
3. **Cell grid** — values, formulas, cached values when available, number formats, comments/notes, hyperlinks.
4. **Ranges and names** — named ranges, tables, merged cells, print areas, data validation, conditional formatting.
5. **Formula graph** — formulas, references, cross-sheet links, external references, volatile functions, circular references where discoverable.
6. **Style/layout signals** — cell fills, font styles, borders, row/column sizes, hidden rows/columns, frozen panes; these are heuristics for module boundaries and user-facing reports.
7. **Controls and assumptions candidates** — hard-coded constants, named inputs, control sheets, data validation lists, scenario selectors.
8. **Outputs candidates** — summary sheets, named outputs, terminal calculated values, charts, pivot-like output tables.
9. **Unsupported/high-risk features** — VBA, macros, unsupported formulas, Solver/Goal Seek, Data Tables, external links, OLE objects, slicers, complex pivots.
10. **Artifacts not yet parsed** — a catch-all bucket with explicit feature type and source location so nothing disappears silently.

If a feature cannot be parsed with existing tools, Phase 1 must record it as unsupported or opaque, not ignore it.

### Coverage invariant

Phase 1 has one hard invariant:

```text
recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects
```

At minimum, every sheet, non-empty cell, formula cell, named range, table, merged range, data validation, conditional format, comment/note, hyperlink, chart, pivot, external link, connection, macro/VBA package part, drawing/OLE/media part, and any detected unsupported package part must appear in either `inventory.json` or `unsupported_or_opaque` diagnostics. Reports must show per-sheet coverage and unclassified/opaque counts.

## Phase 1 module categories

These are Excel-native modules, not final actuarial-domain modules:

| Module category | Meaning | Example evidence |
|---|---|---|
| `workbook_meta` | Global workbook facts and calculation settings | workbook properties, calc mode |
| `sheet_structure` | Sheet-level boundaries and visibility | sheet names, hidden status |
| `input_candidate` | User-editable assumptions or controls | constants, named ranges, validations, colored input cells |
| `data_table` | Rectangular tabular data | Excel tables, dense constant ranges |
| `formula_block` | Connected formula region | dependency graph, copied formula pattern |
| `lookup_block` | Lookup tables and mapping ranges | VLOOKUP/XLOOKUP/INDEX/MATCH references |
| `calculation_chain` | Ordered dependency path to outputs | formula DAG/subgraph |
| `output_candidate` | Results or report outputs | summary sheets, charts, terminal nodes |
| `presentation` | Formatting/report-only content | titles, labels, charts, print areas |
| `external_dependency` | External workbook/data/service dependency | external links, connections |
| `macro_or_code` | VBA/macro/code-bearing elements | `.xlsm`, vba archive presence |
| `unsupported_or_opaque` | Known but not interpreted content | Solver, Data Table, OLE, pivots if unsupported |

A single range can have multiple labels, e.g. `data_table + input_candidate + light_tag:mortality_assumption_candidate`.

## Light actuarial hints only

Phase 1 may attach non-authoritative hints based on names, sheet labels, formulas, and layout:

- `mortality_candidate`
- `lapse_candidate`
- `expense_candidate`
- `premium_candidate`
- `claim_or_benefit_candidate`
- `reserve_candidate`
- `discount_curve_candidate`
- `projection_axis_candidate`
- `scenario_candidate`
- `output_metric_candidate`

These are confidence-scored hints, not final model structure. They must be confirmable and overridable through `confirm/`.

This list is a planning vocabulary, not the runtime enum contract. The implemented `ActuarialHint` values are `assumption`, `rate_table`, `cashflow`, `projection`, `output`, and `unknown`; see [layered architecture](../design/layered_architecture.md). Do not add enum values solely to match this proposal.

## Repository architecture

This plan records the intended Phase 1 sequence, not the complete current runtime map. The legacy `inspect` path runs rule classification, confirmation-question generation, structural completeness checks, numerical cache-versus-recalculation comparison, and handoff. The newer human `step1`/`step2` path performs source conversion/checks, indexing, deterministic views, and evidence navigation. README Step 3's broader semantic analysis and Python generation are not complete. See [layered architecture and artifact boundaries](../design/layered_architecture.md) for implemented files, actual artifact contracts, and planned boundaries.

```text
src/excel_to_act/
  interfaces/    CLI/API/UI adapters; no parsing logic here
  orchestrator/  explicit parse→inventory→graph→classify→confirm→store→report workflows
  plugins/       protocols/registries for replaceable tools
  ingest/        workbook readers; openpyxl first, optional calamine/fastexcel later
  inventory/     complete workbook inventory extraction
  graph/         formula dependency graph, source maps, subgraph selection
  classify/      module classification and light domain hints
  confirm/       human decisions, overrides, conversion scope, thresholds
  store/         persisted artifacts, hashes, lineage, run metadata
  report/        Markdown/HTML/Excel reports
  schemas/       runtime Pydantic artifact contracts
schemas/         exported JSON Schema / contract documentation
examples/        sanitized fixture workbooks and expected outputs
tests/           boundary-level tests per module
```

Each directory owns one capability. Cross-module calls should go through typed artifact contracts, not ad-hoc dictionaries.

## Existing tools to use first

Use existing libraries before rewriting Excel machinery:

- `openpyxl`: default `.xlsx/.xlsm` structure, formulas, styles, names, tables, comments, validations.
- OOXML package/relationship scanning with Python standard `zipfile` + XML metadata: detect workbook parts that openpyxl cannot fully model, such as `xl/vbaProject.bin`, `xl/charts/*`, `xl/pivotTables/*`, `xl/externalLinks/*`, `xl/connections.xml`, drawings, OLE objects, and media.
- `formulas` / `xlcalculator`: optional formula reference parsing/evaluation experiments behind plugin boundaries; Phase 1 graph extraction must not require formula evaluation.
- `networkx`: graph representation and traversal.
- `pydantic`: artifact schemas and validation.
- `xlsxwriter`: later generated report workbooks; Markdown report comes first.
- `LibreOffice` / `xlwings`: future oracle runners, not required for first Phase 1 PRs.

Avoid GPL/AGPL core dependencies in the default path. GPL projects can be research references or isolated optional tools.

## Phase 1 acceptance criteria

By the end of Phase 1, the tool should be able to process a sanitized workbook and produce:

1. `workbook_manifest.json` — file hash, sheets, high-level features.
2. `inventory.json` — cells/ranges/names/tables/styles/validations/links coverage.
3. `dependency_graph.json` — nodes/edges/source locations for formulas and references.
4. `module_classification.json` — Excel-native module categories covering every recognized region plus catch-all opaque artifacts.
5. `confirmation_template.json` — questions/decisions for user confirmation.
6. `phase1_report.md` — readable health report, unsupported features, module breakdown, and recommended next action.

No generated Python model is required for this milestone.

## Non-goals for Phase 1

- No full actuarial object model.
- No automatic production Python generator.
- No performance optimization/vectorization.
- No regulatory sign-off engine.
- No claim that the tool understands all actuarial semantics.
- No silent dropping of unsupported Excel features.

## Implementation tasks

### Task 1: Establish repository skeleton and package metadata

**Objective:** Create a minimal Python package with pluggable directories and docs.

**Files:**
- Create: `pyproject.toml`
- Modify: `README.md`
- Create: `src/excel_to_act/**/README.md`
- Create: `src/excel_to_act/**/__init__.py`

**Verification:**

```bash
python -m compileall src
python -m pip install -e .
excel-to-act --help
```

### Task 2: Define Phase 1 artifact contracts

**Objective:** Add Pydantic models for manifest, inventory, graph, classification, confirmation, and report metadata.

**Files:**
- Create: `src/excel_to_act/schemas/`
- Create/export: `schemas/` JSON Schema docs when useful
- Test: `tests/test_phase1_contracts.py`

**Acceptance:** Contracts serialize to JSON and reject malformed artifacts.

### Task 3: Implement plugin protocols

**Objective:** Define replaceable interfaces for workbook reader, inventory extractor, graph builder, classifier, store, and reporter.

**Files:**
- Create: `src/excel_to_act/plugins/contracts.py`
- Create: `src/excel_to_act/plugins/registry.py`
- Test: `tests/test_plugin_contracts.py`

**Acceptance:** A fake plugin can be registered and called in tests. Plugin inputs/outputs must be Pydantic artifacts. Do not pass openpyxl workbook objects across module boundaries except inside `ingest/` and `inventory/` implementation internals. Keep registry minimal; avoid complex dynamic loading until needed.

### Task 4: Implement openpyxl ingestion

**Objective:** Load workbook files and emit a manifest without losing high-risk feature flags.

**Files:**
- Create: `src/excel_to_act/ingest/openpyxl_reader.py`
- Test: `tests/test_openpyxl_reader.py`

**Acceptance:** Tests cover `.xlsx`, hidden sheets, named ranges, formulas, comments, validations, and `.xlsm` macro detection if fixture is available. Reader also records OOXML package parts that it cannot semantically parse so later inventory can mark them opaque.

### Task 5: Implement core inventory extraction

**Objective:** Extract source-preserving workbook inventory for sheets, cells, names, tables, and ranges.

**Files:**
- Create: `src/excel_to_act/inventory/extractor.py`
- Test: `tests/test_inventory_extractor.py`

**Acceptance:** Every non-empty fixture cell and core workbook object has a `source_location`. Classification is not introduced in this task.

### Task 6: Implement layout and opaque inventory coverage

**Objective:** Add layout/style/validation/protection inventory and explicit unsupported/opaque package-part records.

**Files:**
- Create: `src/excel_to_act/inventory/layout.py`
- Create: `src/excel_to_act/inventory/opaque.py`
- Test: `tests/test_inventory_coverage.py`

**Current plan/code divergence:** These two modules have not been created. The implemented behavior is: layout logic is inline in `inventory/extractor.py`'s `_layout()`; opaque records are in `ingest/ooxml_package.py`'s `OPAQUE_MARKERS` + `scan_ooxml_package()`. This task should **refactor/forward** these two locations into the modules above, rather than write a second implementation and allow further drift.

**Acceptance:** Every non-empty fixture cell and known workbook-level object is represented or explicitly marked unsupported/opaque. Tests must assert the coverage invariant: `recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects`.

### Task 7: Implement graph extraction

**Objective:** Build a formula/reference graph from workbook formulas and named ranges.

**Files:**
- Create: `src/excel_to_act/graph/builder.py`
- Test: `tests/test_graph_builder.py`

**Acceptance:** Cross-sheet references, range references, and unsupported references are captured without crashing. Phase 1 graph extraction parses references and dependencies only: it does not evaluate formulas, does not generate Python expressions, and creates diagnostic nodes for formulas it cannot parse.

### Task 8: Implement Excel-native module classification

**Objective:** Classify ranges/regions into Phase 1 module categories and attach optional light actuarial hints.

**Files:**
- Create: `src/excel_to_act/classify/rules.py`
- Create: `src/excel_to_act/classify/classifier.py`
- Test: `tests/test_classifier.py`

**Acceptance:** Classification covers all recognized regions and produces catch-all records for opaque/unsupported content.

### Task 9: Implement confirmation templates

**Objective:** Generate user-reviewable questions and defaults from classifications.

**Files:**
- Create: `src/excel_to_act/confirm/templates.py`
- Test: `tests/test_confirmation_templates.py`

**Acceptance:** Inputs, outputs, unsupported features, and uncertain module boundaries generate explicit confirmation items.

### Task 10: Implement local artifact store

**Objective:** Persist Phase 1 artifacts with workbook hash and run metadata.

**Files:**
- Create: `src/excel_to_act/store/local_store.py`
- Test: `tests/test_local_store.py`

**Acceptance:** Re-running on the same workbook creates deterministic artifact paths and does not overwrite prior runs unless explicitly requested.

### Task 11: Implement CLI JSON workflow

**Objective:** Provide an end-to-end `excel-to-act inspect workbook.xlsx --out artifacts/` command that writes JSON artifacts.

**Files:**
- Modify: `src/excel_to_act/interfaces/cli.py`
- Create: `src/excel_to_act/orchestrator/phase1.py`
- Test: `tests/test_phase1_cli.py`

**Acceptance:** The CLI writes manifest, inventory, dependency graph, module classification, and confirmation template JSON artifacts.

### Task 12: Implement Markdown report generator

**Objective:** Generate a readable Phase 1 report from persisted artifacts.

**Files:**
- Create: `src/excel_to_act/report/markdown.py`
- Test: `tests/test_phase1_report.py`

**Acceptance:** The report identifies unsupported features, module coverage, per-sheet unclassified/opaque counts, and confirmation checklist items.

## Mapping to the README pipeline

The README divides the long-term roadmap into Steps 0–5. This plan's Phase 1 **targets Steps 0–3** (ingest → inventory → graph → classify → confirm → store → report); it does not mean the broader README Step 3 semantic analysis is already implemented. The legacy `inspect` path has rule classification, confirmation questions, and #7's pre-generation saved-cache versus actual-recalculation check, while human Step 1/2 focuses on decomposition and evidence navigation. The README's Step 4 (Python generation) is excluded by this plan's "Non-goals" section; post-generation code-equivalence remains later work.

### Current artifact layout vs. the Step 1 target

- **Current legacy `inspect` layout:** The CLI command is `excel-to-act inspect <workbook> --out <dir>` and remains single-workbook. It writes flat artifacts under `<out>/workbooks/<sha256>/<run_id>/`, including `workbook_manifest.json`, `inventory.json`, `dependency_graph.json`, `module_classification.json`, `confirmation_template.json`, `completeness.json`, `handoff.json`, `handoff.md`, and `run_metadata.json`, plus artifact-index files and latest-run aliases.
- **Current human Step 1/2 layout:** `excel-to-act step1 convert <directory> --out <dir>` writes source-bound facts, inventory, object/part ledgers, quality and handoff artifacts, plus control/VBA handoffs where available; original source evidence and copied package-part bytes are stored alongside them. Human Step 2 writes a separate batch index and prepared reading package with deterministic views and evidence handoffs. See README §1 and [layered architecture](../design/layered_architecture.md).
- **Step 1 target layout:** See README §3.2 — directories by content type, both Markdown and JSON for each type, and separate contracts for `coverage` and `handoff`. This remains a future layout. Keep `<sha256>/<run_id>/` as the namespace; **do not use a filename slug** (different content with the same name can collide, and renaming breaks links).

### Confirmed hard-constraint gaps

- **The legacy coverage summary is not independent:** `inventory/extractor.py` sets `CoverageSummary.discovered_workbook_objects` directly to `recognized + opaque`, which is an identity and therefore always passes. Human Step 1 has separate source-side coverage checks; #5 still defines the outstanding object-accounting scope (see [layered architecture](../design/layered_architecture.md)).
- **`OPAQUE_MARKERS` has been completed:** it grew from 8 to 18 tokens, adding `xl/model/`, `_xmlsignatures/`, `calcChain.xml`, `/ctrlProps/`, `/slicers/`, `/timelines/`, `/macrosheets/`, `/customXml/`, `volatileDependencies.xml`, and `/webExtensions/`. These parts were previously neither collected nor marked opaque, so they were silently dropped.
- **Encrypted/damaged files raise uncaught exceptions:** `ZipFile()` in `scan_ooxml_package()` has no exception handling, violating the "report an error, do not panic" requirement (issue #8, P1).

## PR planning note

Use `docs/plans/pr_plan_phase1.md` for the proposed PR split. PR-01 should only establish the skeleton and plan so later implementation PRs stay small and reviewable.
