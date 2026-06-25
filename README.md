# excel_to_act

Excel-to-actuarial-model tooling, starting with **Phase 1: lossless Excel decomposition**.

The first milestone does **not** generate Python, decide the final actuarial object model, or rewrite every workbook into a polished software model. It builds a structured, pluggable framework that can decompose an Excel workbook into complete artifacts: workbook inventory, dependency graph, technical modules, light domain tags, human confirmations, storage records, and reports.

Unsupported Excel features are recorded as unsupported/opaque artifacts, not ignored.

See:

- `docs/plans/phase1_excel_decomposition_plan.md`
- `docs/plans/pr_plan_phase1.md`

## Initial package layout

```text
src/excel_to_act/
  interfaces/    CLI/API/UI entrypoints only
  orchestrator/  explicit state-machine workflows
  plugins/       replaceable tool contracts and registries
  ingest/        workbook readers, starting with openpyxl
  inventory/     complete workbook inventory extraction
  graph/         formula/dependency graph and source maps
  classify/      Excel technical classification + light tags
  confirm/       human confirmation decisions
  store/         artifact persistence and lineage
  report/        health/module/audit reports
  schemas/       runtime Pydantic artifact contracts
schemas/         exported JSON Schema / contract docs
examples/        sanitized fixture workbooks and expected outputs
tests/           boundary-focused tests
```

## Phase 1 boundaries

Phase 1 stops at workbook decomposition and reviewable artifacts:

- preserve original workbook hierarchy and source locations;
- classify content into Excel-native modules;
- create confirmation templates for uncertain inputs/outputs/module boundaries;
- persist artifact lineage for later generation and audit.

Later phases can generate faithful Python, actuarial-domain Python, APIs, validation packages, and governance packages once the decomposition layer is trusted.
