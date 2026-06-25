# Phase 1 PR Plan

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** Deliver the Excel decomposition foundation through small PRs, stopping before final actuarial-domain restructuring or Python model generation.

**Architecture:** Each PR adds one pluggable boundary or one artifact stage. The repo remains usable after every PR, with tests proving the new boundary and preserving original workbook source locations.

**Tech Stack:** Python 3.11+, openpyxl, Python stdlib OOXML zip/XML scanning, Pydantic, NetworkX, Typer, pytest, optional formulas/xlcalculator behind plugin interfaces.

---

## Current confirmation needed

The final actuarial structure is intentionally **not** fixed yet. Phase 1 should only decompose Excel according to the workbook's own structure and classify all content into Excel-native modules, with light confirmable actuarial hints.

Open product question for later phases:

> What canonical actuarial objects should Phase 3 use: product/policy/model point/assumption/scenario/projection/cashflow/output, or a lighter generic calculation model first?

This PR plan avoids depending on that answer.

## Hard Phase 1 invariants

1. Preserve original hierarchy: `workbook → sheet order → sheet → row/column → cell/range/object → source_location`.
2. Classification is an overlay; it must not replace raw workbook inventory.
3. Coverage equation must be testable: `recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects`.
4. Unsupported Excel features are recorded, not ignored.
5. Formula graph work parses dependencies only; no formula evaluation and no Python generation in Phase 1.
6. Plugin stages exchange typed artifact contracts, not raw openpyxl objects across the whole pipeline.

## PR overview

| PR | Goal | Main files | Depends on | Acceptance |
|---|---|---|---|---|
| PR-01 | Skeleton + architecture plan | `pyproject.toml`, `README.md`, `src/excel_to_act/*/README.md`, `docs/plans/*` | none | package imports/compiles; plan reviewed |
| PR-02 | Runtime artifact contracts | `src/excel_to_act/schemas/`, `schemas/`, tests | PR-01 | contracts serialize/validate |
| PR-03 | Minimal plugin interfaces | `plugins/contracts.py`, `plugins/registry.py`, tests | PR-02 | fake plugins register/call via typed artifacts |
| PR-04 | openpyxl + OOXML package ingestion | `ingest/openpyxl_reader.py`, `ingest/ooxml_package.py`, fixtures, tests | PR-03 | manifest extracts workbook facts and opaque package parts |
| PR-05 | Core inventory extraction | `inventory/extractor.py`, tests | PR-04 | sheets/cells/names/tables/ranges represented |
| PR-06 | Layout + opaque inventory coverage | `inventory/layout.py`, `inventory/opaque.py`, tests | PR-05 | styles/layout/validations/protection/unsupported parts covered |
| PR-07 | Formula/source graph | `graph/builder.py`, tests | PR-06 | formulas and references become graph nodes/edges without evaluation |
| PR-08 | Module classification | `classify/rules.py`, `classify/classifier.py`, tests | PR-07 | every recognized region gets category or catch-all |
| PR-09 | Human confirmation templates | `confirm/templates.py`, tests | PR-08 | uncertain decisions become user-reviewable items |
| PR-10 | Local artifact store | `store/local_store.py`, tests | PR-09 | run artifacts persisted with hashes/metadata |
| PR-11 | Orchestrator + CLI JSON artifacts | `orchestrator/phase1.py`, `interfaces/cli.py`, tests | PR-10 | CLI writes JSON artifacts |
| PR-12 | Markdown report | `report/markdown.py`, tests | PR-11 | report shows coverage, opaque content, and confirmation checklist |

## PR-01: Skeleton + architecture plan

**Goal:** Establish a pluggable directory framework and document the Phase 1 direction.

**Scope:**

- Minimal Python package metadata.
- Directory-per-capability structure.
- README explaining Phase 1 lossless decomposition.
- Plan docs with Phase 1 boundaries and future PR split.
- CLI placeholder only; no real workbook parsing yet.

**Suggested files:**

- `pyproject.toml`
- `README.md`
- `docs/plans/phase1_excel_decomposition_plan.md`
- `docs/plans/pr_plan_phase1.md`
- `src/excel_to_act/interfaces/`
- `src/excel_to_act/orchestrator/`
- `src/excel_to_act/plugins/`
- `src/excel_to_act/ingest/`
- `src/excel_to_act/inventory/`
- `src/excel_to_act/graph/`
- `src/excel_to_act/classify/`
- `src/excel_to_act/confirm/`
- `src/excel_to_act/store/`
- `src/excel_to_act/report/`
- `src/excel_to_act/schemas/`
- `schemas/`, `examples/`, `tests/`

**Acceptance criteria:**

- `python -m compileall src` passes.
- `python -m pip install -e .` succeeds in a clean venv.
- `excel-to-act --help` works.
- A reviewer agrees the plan does not prematurely force Phase 3 actuarial structure.

## PR-02: Runtime artifact contracts

**Goal:** Define stable schemas before implementing behavior.

**Scope:**

- Runtime Pydantic models under `src/excel_to_act/schemas/`.
- Optional exported JSON Schema docs under top-level `schemas/`.
- `WorkbookManifest`, `WorkbookInventory`, `FormulaGraph`, `ModuleClassification`, `ConfirmationTemplate`, `RunMetadata`, `UnsupportedFeature`, and `SourceLocation`.

**Acceptance criteria:**

- Tests prove JSON roundtrip.
- Tests reject malformed source locations and unknown required enum values.
- Schema version appears in every artifact.

## PR-03: Minimal plugin interfaces

**Goal:** Make ingestion/classification/reporting replaceable without overengineering dynamic loading.

**Scope:**

- Protocols for reader, inventory extractor, graph builder, classifier, store, reporter.
- Simple in-process registry keyed by plugin name and capability.
- Rule: plugin inputs/outputs are Pydantic artifacts.

**Acceptance criteria:**

- Fake plugin test proves registration and dispatch.
- No concrete implementation leaks into interface tests.
- openpyxl workbook objects do not cross module boundaries except inside reader/inventory internals.

## PR-04: openpyxl + OOXML package ingestion

**Goal:** Use existing tools for Excel reading and explicitly detect content openpyxl cannot fully parse.

**Scope:**

- Load `.xlsx/.xlsm` with openpyxl.
- Emit manifest: sheets, hidden state, dimensions, calc settings, named ranges count, macro flag, file hash.
- Add package scanner using `zipfile` / OOXML relationships for `xl/vbaProject.bin`, charts, pivots, external links, connections, drawings, OLE, media.

**Acceptance criteria:**

- Fixture workbook tests cover visible/hidden sheets, formulas, named ranges.
- Unsupported file type returns a typed error.
- Opaque package parts are recorded in manifest diagnostics.

## PR-05: Core inventory extraction

**Goal:** Represent source-preserving core workbook content.

**Scope:**

- Sheets, sheet order, dimensions.
- Non-empty cells, formulas, cached/value fields where available, number formats.
- Named ranges, tables, merged ranges.
- Source locations for every object.

**Acceptance criteria:**

- Test fixture expected counts match inventory counts.
- Every non-empty cell has `source_location`.
- Classification is not introduced in this PR.

## PR-06: Layout + opaque inventory coverage

**Goal:** Extend inventory so Phase 1 can claim coverage of non-formula Excel content.

**Scope:**

- Styles/layout signals: fills, fonts summary, row/column sizes, hidden rows/columns, frozen panes.
- Data validations, conditional formatting, comments, hyperlinks, protection.
- Unsupported/opaque package parts discovered in PR-04.

**Acceptance criteria:**

- Known unsupported features are recorded, not ignored.
- Test asserts `recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects` for fixtures.
- Per-sheet inventory counts can be computed.

## PR-07: Formula/source graph

**Goal:** Convert formula references into a graph suitable for later dependency slicing.

**Scope:**

- Node model for cells, ranges, names, external refs, unsupported refs.
- Edge model for references.
- Graph serialization.
- Optional formula parser adapters behind plugin boundary.

**Acceptance criteria:**

- Cross-sheet formula fixture produces expected edges.
- Range references are preserved even if not expanded to every cell.
- Unsupported formulas produce diagnostic nodes.
- No formula evaluation and no Python expression/code generation.

## PR-08: Module classification

**Goal:** Split workbook content into Excel-native modules that cover the workbook.

**Scope:**

- Rules for input candidates, data tables, formula blocks, lookup blocks, outputs, presentation, external dependencies, unsupported/opaque.
- Light confidence-scored actuarial hints only.
- Classification reasons and source artifact references.

**Acceptance criteria:**

- Every recognized range/region has at least one category or catch-all.
- Classification reasons are recorded for auditability.
- Raw inventory remains unchanged.
- No test asserts final actuarial object structure.

## PR-09: Human confirmation templates

**Goal:** Create the confirmation contract for future UI/API/Hermes review flows.

**Scope:**

- Questions for uncertain inputs/outputs/module boundaries.
- Questions for unsupported/high-risk features.
- Decision records with reviewer, timestamp, source artifact version.

**Acceptance criteria:**

- Confirmation template generated from sample classification.
- User override can change category/hint without mutating raw inventory.
- No new `ai_interface` dependency is introduced; `confirm/` owns the contract.

## PR-10: Local artifact store

**Goal:** Persist every artifact with lineage and reproducibility.

**Scope:**

- Local filesystem store.
- Workbook hash/run id paths.
- Artifact index.

**Acceptance criteria:**

- Same workbook hash writes under deterministic namespace.
- Multiple runs are preserved.
- Store can read back each artifact and validate schema version.

## PR-11: Orchestrator + CLI JSON artifacts

**Goal:** Make Phase 1 pipeline executable from one command without report formatting complexity.

**Scope:**

- `excel-to-act inspect workbook.xlsx --out artifacts/`
- State-machine orchestration.
- JSON artifact writes.

**Acceptance criteria:**

- CLI writes:
  - `workbook_manifest.json`
  - `inventory.json`
  - `dependency_graph.json`
  - `module_classification.json`
  - `confirmation_template.json`
- End-to-end fixture test passes.

## PR-12: Markdown report

**Goal:** Add a readable Phase 1 report after JSON artifacts are stable.

**Scope:**

- `phase1_report.md` generator.
- Per-sheet coverage table.
- Unsupported/opaque content section.
- Confirmation checklist summary.

**Acceptance criteria:**

- Report identifies unsupported features and module coverage.
- Report shows per-sheet unclassified/opaque counts.
- Markdown snapshot test passes for fixture workbook.

## Risks and mitigations

| Risk | Mitigation |
|---|---|
| Excel features exceed openpyxl coverage | Use OOXML package scanner + explicit `unsupported_or_opaque` artifacts |
| Premature actuarial abstraction | Keep Phase 1 Excel-native; only confidence-scored hints |
| Formula parser lock-in | Hide formulas/xlcalculator behind plugin boundary |
| PRs become too large | Split inventory and final CLI/report work into separate PRs |
| Large workbook performance | Phase 1 correctness first; add streaming/fast readers later |
| Client workbook confidentiality | Only sanitized fixtures in repo; no real workbooks committed |
| GPL dependency contamination | Keep GPL tools out of default dependency path |

## Implementation order

1. Merge PR-01 first to establish shared structure and plan.
2. Implement schemas before concrete readers.
3. Add minimal plugin contracts before choosing formula engines.
4. Build ingestion/inventory/graph/classification in order.
5. Add confirmation/store/orchestrator/report only after artifact contracts stabilize.
6. Start Python generation only in a later separate plan after Phase 1 reports prove decomposition coverage.
