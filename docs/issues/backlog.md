# Issue Backlog

Reconciled with live GitHub snapshots and merged PR evidence on October 8, 2026. GitHub Issue #3 is the Epic; its child Issues are #4–#14. Preserve the scope and dependency edges in each live Issue body.

This PR aligns the backlog and closes Epic #3 once merged. All child Issues #4–#14 are closed. Delivery: #4 PRs #15/#16; #5 #38; #6 #27; #7 #41; #8 #37; #9 #39 plus #40; #10 deferred with parser-fidelity research incomplete; #11 #36; #12 folded into #14; #13 #28; #14 #29, including the #12 contract. See each item for limits and evidence.

<!-- ISSUE_NUMBER: 3 -->
<!-- ISSUE: docs: Phase 1.5 actual tasks, dependencies, and phase acceptance (Epic) -->

## Issue 3 docs: Phase 1.5 actual tasks, dependencies, and phase acceptance (Epic)

[GitHub Issue #3](https://github.com/ferryhe/excel_to_act/issues/3) · Phase 1.5 Epic closeout

### Current status

This PR aligns the backlog with live child Issue statuses, scope, dependencies, and merged PR evidence, and closes the Epic once merged. The closeout records #10 as an accepted deferral with parser-fidelity research incomplete; #12's requirements are consolidated into #14 with no separate implementation. The original scope and dependency ordering remain unchanged.

### What to do

Make the decomposition artifacts reliable first, then proceed with Agent views and numerical verification. Implementation details and rules belong in the child Issues; this Epic does not maintain a second design.

### Actual child items

| Issue | Status | Priority | Current work |
|---|---|---|---|
| [#4](https://github.com/ferryhe/excel_to_act/issues/4) | Closed; accepted | Complete | Cache ingestion and boundary regression; PRs #15/#16 |
| [#8](https://github.com/ferryhe/excel_to_act/issues/8) | Closed | P0 | Controlled read-failure diagnostics and CLI path; PR #37 includes real encrypted and damaged inputs |
| [#5](https://github.com/ferryhe/excel_to_act/issues/5) | Closed | P0 | Identity-based coverage, per-sheet/per-kind gaps, failure handoff and nonzero exit; PR #38 |
| [#13](https://github.com/ferryhe/excel_to_act/issues/13) | Closed | P0 | Fidelity rules, raw fields, schema and read-back; PR #28 |
| [#6](https://github.com/ferryhe/excel_to_act/issues/6) | Closed | P0 | Tokenizer-based reference graph; PR #27 |
| [#11](https://github.com/ferryhe/excel_to_act/issues/11) | Closed | P1 | Layer boundaries and artifact contracts; PR #36 |
| [#14](https://github.com/ferryhe/excel_to_act/issues/14) | Closed; includes #12 | P1 | Deterministic views and reference validation, including #12 contract; PR #29 |
| [#9](https://github.com/ferryhe/excel_to_act/issues/9) | Closed | P1 | Default-runtime license gate and selected formulas fixture; PRs #39/#40 |
| [#7](https://github.com/ferryhe/excel_to_act/issues/7) | Closed | P1 | Legacy `inspect` cache comparison; PR #41, with incomplete cases documented |
| [#10](https://github.com/ferryhe/excel_to_act/issues/10) | Closed as deferred | Reopen as needed | Accepted deferral; parser-fidelity research was not performed and remains incomplete |
| [#12](https://github.com/ferryhe/excel_to_act/issues/12) | Closed; folded into #14 | With #14 | No separate implementation; requirements delivered under #14 / PR #29 |

The items #4–#14 above are the child Issues in this Epic. Four historical entries numbered #15–#18 were local implementation records, now labeled `LOCAL-*`; they were never GitHub Issue IDs. GitHub #15–#17 are PRs, and GitHub #18 is a separate real Step 2 Issue outside this Epic. Keep the local records for data tables, VML controls, VBA, and completeness/handoff, with PR #15 as evidence. PR #15 does not prove every semantic feature or real VBA integration path is complete.

### Phases and dependencies

1. Align the backlog: this documentation PR synchronizes real IDs, scope, status, local records, and README, and closes Epic #3 once merged. #4 is accepted; #10 is deferred with parser-fidelity research incomplete; #12 is folded into #14.
2. Decomposition dependencies: #5 depends on #11; #13 depends on #11 and aligns with #5 by object responsibility; #6 depends on #11/#13 and final integration uses #5. #8 is independent.
3. Enable Agent use: #11 delivered the architecture document, and #14 delivered the #12 contract.
4. Numerical verification: #9 checked the selected backend and default dependency policy; #7 delivered one real comparison path.

These dependency edges preserve the live child Issue bodies. #8/#9 were independent; #7 did not depend on views. Child closure does not mean capabilities exceed the limits recorded below.

### How to verify

- [x] Every item has a real link and accurate status, scope, and dependencies; the README, local backlog, and GitHub agree; local tasks in the README are no longer labeled issues #15/#16/#17.
- [x] #4 cache-boundary regression has been accepted and closed; PR #15/#16 are merged, and the implementation boundaries and evidence for data tables/controls/VBA/handoff are retained.
- [x] Invalid input produces readable failure artifacts and stops downstream work; valid input can be read back.
- [x] Omissions fail under unified object accounting, raw facts remain traceable, and the reference graph does not create false edges.
- [x] Minimal views compile deterministically, retain all records within budget, and allow Agent source references to be validated.
- [x] Numerical verification has a real independent result; missing sources are clearly marked unverified, and differences include location and tolerance.
- [x] Close the Epic only after its child items are complete; creating child Issues does not mean implementation is complete.

For this closeout, #10 is an accepted deferral and parser-fidelity research remains incomplete. #12 was consolidated into #14; it had no separate implementation.

Create directories when their documents are ready; do not add empty scaffolding. Record tests using real encrypted files, real VBA integration samples, and recalculation backends under their respective paths; existing unit tests do not replace unperformed integration acceptance.

---

<!-- ISSUE_NUMBER: 4 -->
<!-- ISSUE: feat(ingest): cached_value ingestion and boundary regression acceptance complete -->

## Issue 4 feat(ingest): cached_value ingestion and boundary regression acceptance complete

[GitHub Issue #4](https://github.com/ferryhe/excel_to_act/issues/4) · Closed: accepted; PRs #15/#16 merged

### Current status

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`.

The remaining cache-boundary regression was merged into main by [PR #16](https://github.com/ferryhe/excel_to_act/pull/16) on 2026-10-03 at 15:08 UTC (squash commit `f12871b`), which closed this Issue via `Closes #4`. [PR #15](https://github.com/ferryhe/excel_to_act/pull/15) remains the historical delivery evidence for the core cache ingestion capability.

### What to do

Retain the formula text, cached result, and a flag indicating whether the cache exists for the same formula cell. Recalculation, cache freshness, and numerical comparison belong to [#7](https://github.com/ferryhe/excel_to_act/issues/7); raw cache text, date systems, and formula-property fidelity belong to [#13](https://github.com/ferryhe/excel_to_act/issues/13).

### Delivered work and implementation boundaries

- [src/excel_to_act/ingest/cached_values.py](https://github.com/ferryhe/excel_to_act/blob/c86fa5585b7a8ca0d27642f9f958379d25c120fb/src/excel_to_act/ingest/cached_values.py): reads the data-only values and merges them by sheet and coordinate.
- [src/excel_to_act/inventory/extractor.py](https://github.com/ferryhe/excel_to_act/blob/c86fa5585b7a8ca0d27642f9f958379d25c120fb/src/excel_to_act/inventory/extractor.py): merges cached values, does not mark constants as formula caches, and records one warning when formulas throughout the workbook have no cache.
- [CellInventory](https://github.com/ferryhe/excel_to_act/blob/c86fa5585b7a8ca0d27642f9f958379d25c120fb/src/excel_to_act/schemas/artifacts.py) and [schemas/workbook_inventory.schema.json](https://github.com/ferryhe/excel_to_act/blob/c86fa5585b7a8ca0d27642f9f958379d25c120fb/schemas/workbook_inventory.schema.json): `cached_value`, `cached_value_available`.
- [tests/test_cached_value.py](https://github.com/ferryhe/excel_to_act/blob/c86fa5585b7a8ca0d27642f9f958379d25c120fb/tests/test_cached_value.py): tests for cached values, missing caches, constants, and avoiding duplicate counts across two loads.
- The inventory README is aligned with the implementation.

The original requirement to read caches in `openpyxl_reader.py` was changed to use `cached_values.py` as the single source of truth: the manifest does not contain cells, so reading them all again in the reader would add I/O without providing a place to store the results.

### How to verify

- [x] A formula cell with a cache retains both `formula` and `cached_value`, with `available=True`.
- [x] A constant cell has `cached_value=None` and `available=False`.
- [x] When formulas throughout a workbook have no stored results, `available=False` and a `missing_cached_values` warning is recorded.
- [x] Loading the same cell twice does not count it twice.
- [x] PR #16: all 5 cache-specific tests and all 26 tests passed; Ruff, compileall, and `excel-to-act --help` passed.
- [x] Automated tests cover 0, False, and the empty string all having `available=True`; typed-string `<v/>` can be distinguished from uncached untyped `<v/>`.
- [x] PR #15's core implementation was merged into main (`b7270f5`); the fields and basic behavior remain consistent.
- [x] GitHub Actions passed on Python 3.11, 3.12, and 3.13.
- [x] The fields and behavior were reviewed after PR #16 was merged, and Issue #4 was closed.

Cache presence is determined by `available`; automated regression tests now confirm that valid values such as 0, False, and the empty string are not treated as missing just because they are empty. This change completes boundary acceptance without changing the core ingestion implementation. Raw-text/type fidelity and missing-cache comparison policy are accepted under #13 and #7, respectively, and are not reimplemented in this Issue.

---

<!-- ISSUE_NUMBER: 5 -->
<!-- ISSUE: fix(coverage): unify object accounting and block downstream use when objects are missed -->

## Issue 5 fix(coverage): unify object accounting and block downstream use when objects are missed

[GitHub Issue #5](https://github.com/ferryhe/excel_to_act/issues/5) · Closed; PR #38 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P0.

**Current delivery:** Closed by [PR #38](https://github.com/ferryhe/excel_to_act/pull/38). `CoverageSummary.discovered_workbook_objects` comes from `len(scan["objects"])`, and recognized count is based on matched source identities. Summary counts alone do not prove identity-by-identity coverage; `verify_completeness` compares source and inventory identities, validates stored counts, reports per-sheet/per-kind gaps, saves failure evidence, and exits nonzero for omissions. Human `step1` has separate logical-object and package-part ledgers.

At the October 3 planning baseline, the inventory totals were not checked by identity and the earlier completeness checks had the reproduced limits described here. Changing the then-info-level `coverage_arithmetic` severity alone could not provide the independent comparison added in PR #38.

### What to do

Independently discover and account for outputs for the currently declared objects, with unified count units, stable identities, and clear boundaries for unresolved items. Do not reimplement the existing completeness module.

### Deliverables and scope

- An accounting table for currently covered objects: sheets, non-empty cells, names, Tables, merged ranges, data validations, conditional formatting, comments, hyperlinks, existing layout objects, and the data tables, controls, and VBA parts in the current branch.
- An independent OOXML discovery path and object identities; document logical-object coverage separately from package-part accounting, without counting a part and its recognized internal objects twice.
- Count only real unresolved objects with `opaque=True` as opaque; warnings such as `missing_cached_values` are not objects, and multiple diagnostics must not count one object more than once.
- Match corresponding outputs to source objects by identity; do not use global counts by type or consume arbitrary evidence as a substitute.
- Reuse the existing CompletenessReport, report/handoff.py, and store, and output expected/actual/gaps for each sheet and object type.
- On coverage imbalance, still save failure diagnostics and handoff; move CoverageSummary validation so a model exception does not abort before saving.
- Export new fields to JSON Schema, explain how old artifacts are read back, and add necessary regression tests for missed objects.

Out of scope: interpreting VBA/chart/pivot content; a new report-template system.

### How to verify

- [ ] A normal fixture balances under unified accounting.
- [ ] Deliberately skipping declared objects such as cells, names, conditional formatting, or Tables causes accounting to fail and identifies the source.
- [ ] Missing parts, XML scan failures, and zero objects are handled separately; missing evidence cannot pass.
- [ ] Non-opaque warnings do not change coverage, and multiple diagnostics do not count one object more than once.
- [ ] Output the three counts and detailed differences per sheet and object type; identify the source of package-level opaque objects separately without inventing a sheet.
- [ ] A real omission becomes an error, and the CLI returns nonzero after saving a failed handoff.
- [ ] Schema, store read-back, existing tests, and Ruff pass.

Dependency: the minimal responsibility/object conventions in [#11](https://github.com/ferryhe/excel_to_act/issues/11); do not wait for the full architecture document or a nonexistent report/markdown.py.

---

<!-- ISSUE_NUMBER: 6 -->
<!-- ISSUE: fix(graph): use tokenizer and minimal reference parsing to build reliable dependency edges -->

## Issue 6 fix(graph): use tokenizer and minimal reference parsing to build reliable dependency edges

[GitHub Issue #6](https://github.com/ferryhe/excel_to_act/issues/6) · Closed; PR #27 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P0.

**Current delivery:** [PR #27](https://github.com/ferryhe/excel_to_act/pull/27) replaced regex scanning with tokenizer operands for cell/range, scoped name, structured-table, and external-workbook references. Dynamic and unsupported references remain outside this scope.

At the October 3 planning baseline, `graph/builder.py` used REF_RE/EXTERNAL_RE and had the reproduced issues listed here; PR #27 delivered the tokenizer-based path summarized above.

GraphNodeKind.name already exists and is no longer listed as a type to add.

### What to do

Use the tokenizer from the current openpyxl dependency to distinguish tokens, then minimally parse reference-type operands. The tokenizer does not parse Table columns, name scope, or external workbook identity, so add the necessary resolution for those cases.

### Deliverables and scope

- In graph/builder.py: parse reference tokens, handle quoted sheet names correctly, and preserve external identities.
- Reuse existing cell/range/name/external nodes; represent structured references with range plus metadata where possible, and add a type only when they cannot otherwise be represented.
- Collect workbook/sheet name scope and necessary Table-column facts along the real reference-resolution path; do not build a parallel graph.
- Preserve source_location and unresolved records for references that cannot be located statically; parsing failures must not crash.
- Keep orchestrator, VBA in the same node namespace, and the verify/formulas_linked call contract in sync.
- Summarize tokenizer coverage in tests/test_graph_builder.py.

Out of scope: formula evaluation, a full AST, function semantics, and guesses about dynamic references. Reference parsing must not add EUPL/GPL dependencies.

### How to verify

- [ ] `="A1"` produces no reference edge and is not recorded as a parsing omission for a valid formula with zero references.
- [ ] `=1+1` is a valid formula with zero dependencies; completeness does not report an error just because there are no outgoing edges.
- [ ] `=Table1[Col]` produces a traceable structured reference and is not mistaken for an external file.
- [ ] `=SUM(MortRate)` produces a name node, and same-name sheet shadowing works correctly.
- [ ] `=Assumptions!B7` and `='O''Brien'!$B$2` preserve the complete sheet name.
- [ ] `=[Book.xlsx]Inputs!A1` preserves external workbook identity and does not produce an equivalent false local cell edge.
- [ ] References that cannot be parsed or located statically have clear diagnostics and original locations, without crashing.
- [ ] Existing VBA integration and node namespaces remain consistent; tests and Ruff pass.

Dependencies: the required source and field conventions in #11/#13; final integration uses #5 coverage checks. This item does not wait for views or a numerical oracle.

---

<!-- ISSUE_NUMBER: 7 -->
<!-- ISSUE: feat(validation): compare cached baselines with one optional recalculation source -->

## Issue 7 feat(validation): compare cached baselines with one optional recalculation source

[GitHub Issue #7](https://github.com/ferryhe/excel_to_act/issues/7) · Closed; PR #41 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P1.

**Current delivery:** [PR #41](https://github.com/ferryhe/excel_to_act/pull/41) added the legacy `inspect` comparison against independent `formulas` recalculation, with coverage, difference locations, and tolerance reporting. Unsupported 1904-date comparisons and unknown/unsupported formulas remain incomplete; an unavailable optional backend is `not_run`. This is not complete Excel recalculation.

Cache ingestion and boundary regression for [#4](https://github.com/ferryhe/excel_to_act/issues/4) were merged into main in PR #15/#16 and accepted. The legacy `inspect` path now has `verify/numerical.py` and `ValidationReport`: it compares saved caches with optional, actual `formulas` recalculation and records coverage gaps. The separate human Step 1/2 and post-generation workflows do not consume this report yet. Copying a cache and comparing it with the same cache does not count as independent numerical verification.

### What to do

Build the first real numerical-verification loop: use the workbook's saved cache as the baseline, add one actual recalculation source, and compare values by sheet/address.

### Deliverables and scope

- ValidationReport: source, coverage, result status, tolerance, and detailed differences; serializable and stored in the existing store/handoff.
- A minimal coordinate-based comparison function and one optional recalculation adapter in validation/; the protocol covers only what actual calls require.
- Clearly distinguish structural CompletenessReport from numerical ValidationReport; structural pass does not mean numerical pass.
- First verify installation and small-sample capability of the existing oracle-formulas through #9; if it cannot support the initial fixture, choose LibreOffice and record why. Deliver only one verified backend in the first release.
- Explain actual verification coverage when caches are missing, freshness is unknown, the backend is unavailable, or formulas are unsupported; do not present these cases as passing.
- Align the original PR plan to this loop and add numerical-report read-back and real-comparison tests.

Out of scope for the first release: a second recalculation backend, Excel COM, actuarial semantic correctness, and Python generation. Do not build every adapter in parallel without a clear coverage requirement.

### How to verify

- [ ] The same fixture has both a cached baseline and an actual recalculation result, with their distinct sources retained.
- [ ] An intentionally introduced difference beyond tolerance fails and identifies the sheet, address, baseline, actual value, and tolerance.
- [ ] Comparison rules are explicit for numbers, dates, booleans, and Excel error values.
- [ ] Missing caches or an unavailable backend are marked as not run/incomplete, not as verified and passing.
- [ ] Default CI verifies fallback behavior when the backend is not installed; a separate real integration check installs the backend and calculates the fixture, rather than skipping everything.
- [ ] ValidationReport can be stored and read back, and handoff distinguishes structural from numerical status.
- [ ] Tests and Ruff pass; CI does not depend on Excel COM.

Dependencies: #4 cache fields and boundary acceptance (complete, PR #15/#16), #13 value-type and fidelity conventions, and #9 compatibility checks for the selected backend. Independent of #14 views; their order can be adjusted for the next milestone.

---

<!-- ISSUE_NUMBER: 8 -->
<!-- ISSUE: fix(ingest): save diagnostics on read failure and let the CLI exit cleanly -->

## Issue 8 fix(ingest): save diagnostics on read failure and let the CLI exit cleanly

[GitHub Issue #8](https://github.com/ferryhe/excel_to_act/issues/8) · Closed; PR #37 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P0.

**Current delivery:** [PR #37](https://github.com/ferryhe/excel_to_act/pull/37) saves diagnostic-only failed runs, stops downstream extraction, and reports controlled CLI errors. The saved `issue-8/` evidence includes real encrypted and damaged inputs. The change does not decrypt files or add unsupported-format parsing.

At the October 3 planning baseline, ZIP/read failures could escape and downstream extraction was not stopped. PR #37 closed that path and its `issue-8/` evidence includes real encrypted and damaged inputs.

### What to do

Fix the complete read-failure path from reader/scanner → orchestrator → store/handoff → CLI.

### Deliverables and scope

- reader/scanner catches expected file, ZIP, required-XML, and openpyxl read failures.
- Stop the extractor and subsequent phases for an error manifest; record UnsupportedFeature(severity=error).
- Reuse the existing error-record and artifact mechanisms, and define the minimal artifact set available for a failed run; do not fabricate a complete inventory or missing source data.
- The CLI prints a readable failure reason, saves diagnostic locations, and returns a controlled nonzero exit code.
- tests/test_ingest_robustness.py: end-to-end CLI checks and verification that downstream calls stop.
- Keep conversion/decryption integration notes as a brief summary; they are not prerequisite research for fixing the current failure path.

Out of scope: actual decryption, xlsb/xls parsing, automatic conversion, generic retries, or an error framework.

### How to verify

- [ ] Real encrypted xlsx: no uncaught read exception, an error record is saved, and the CLI exits nonzero.
- [ ] Non-ZIP input, truncated ZIP, and ZIP with corrupted/missing required XML: same behavior.
- [ ] xlsb/xls/csv: explicit file_type error; the extractor is not called afterward.
- [ ] Failed runs retain locatable diagnostics and a failure handoff; missing artifacts are listed explicitly.
- [ ] Tests prove that invalid input does not run downstream graph/classify and other stages.
- [ ] Normal xlsx/xlsm paths and artifact read-back remain valid; tests and Ruff pass.

This item can proceed independently; it does not wait for Docling, an oracle, or a complete architecture design.

---

<!-- ISSUE_NUMBER: 9 -->
<!-- ISSUE: chore(build): check core dependency boundaries and compatibility of the selected oracle -->

## Issue 9 chore(build): check core dependency boundaries and compatibility of the selected oracle

[GitHub Issue #9](https://github.com/ferryhe/excel_to_act/issues/9) · Closed; PRs #39/#40 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P1.

**Current delivery:** PRs [#39](https://github.com/ferryhe/excel_to_act/pull/39) and [#40](https://github.com/ferryhe/excel_to_act/pull/40) enforce the default-runtime dependency license gate and verify the selected formulas opt-in on the fixture. The default installation excludes `formulas`; this does not establish compatibility or license status for every optional extra.

The October 3 planning baseline recorded the extras and Python matrix from PR #15; PRs #39/#40 later delivered the default-runtime license gate and selected formulas fixture check, as summarized above.

### What to do

Address only the remaining dependency boundaries and the practical viability of the first recalculation backend for #7; do not split extras again.

### Deliverables and scope

- Check core and transitive dependencies in a clean default runtime environment; do not include dev tools or explicitly opted-in oracle environments in the core assessment.
- Fix the current core-dependency policy for acceptance under original #9: GPL, AGPL, and EUPL-family licenses must not enter the default runtime dependency closure; normalize license variants by identifier. This is a project dependency-selection policy; record it in a short ADR/note without claiming a legal obligation.
- Automatically check license fields/expressions for the default-installed runtime dependency closure in CI, including actually resolved transitive dependencies; fail when any of the listed licenses are found. Validate the same CI check entry point with a known violating sample.
- Reuse an existing metadata tool or a minimal script; define how to handle mixed licenses and missing/indeterminate metadata rather than silently reporting the check as passed.
- The default installation does not include formulas; only explicit oracle-formulas opt-in introduces it.
- Record the tested Python support matrix; do not arbitrarily add an upper limit just because future versions have not been tested.
- Test installation, import, and fixture calculation for the backend selected by #7; import success alone does not prove calculation works.
- Correct A1's descriptions of the old bundled extra, old CI matrix, and incomplete items.

fastexcel is not currently used and is not an acceptance requirement for this Issue; an unselected backend does not block the first release. Do not expand this item into a general compliance system or legal conclusion about software distribution.

### How to verify

- [x] Extras are separated, core CI includes 3.11/3.12/3.13, and these changes were merged into main in PR #15.
- [ ] A clean default installation does not include formulas; only explicit opt-in introduces it.
- [ ] CI enforces a license gate on the default runtime dependency closure, including transitive dependencies, and fails for GPL/AGPL/EUPL matches; dev tools and explicitly opted-in oracle dependencies are excluded.
- [ ] A known violating sample makes the CI check entry point fail; missing or indeterminate license metadata has a clear documented disposition and cannot silently pass.
- [ ] A short decision record documents the policy's excluded license set, mixed-license handling, and metadata sources, consistent with the gate rules.
- [ ] The selected backend installs, imports, and calculates the small #7 sample on supported declared versions.
- [ ] Measured results, version ranges, and the A1 description are consistent.
- [ ] CI, tests, and Ruff pass.

Dependency: choose the first-release backend together with #7; this does not depend on Agent views. Core dependency checks can proceed earlier.

---

<!-- ISSUE_NUMBER: 10 -->
<!-- ISSUE: docs(research): defer the Docling / markitdown fidelity assessment -->

## Issue 10 docs(research): defer the Docling / markitdown fidelity assessment

[GitHub Issue #10](https://github.com/ferryhe/excel_to_act/issues/10) · Closed: deferred; focused research incomplete

### Disposition

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`.

This item is closed as deferred, not because the research is complete. The current approach uses openpyxl + OOXML, and there is no specific work item to integrate Docling/markitdown; the gaps in #5/#6/#8/#13 have direct fix paths, and a focused comparison does not block them.

### Retained decision record

Include the existing selection rationale in the short architecture decision note for [#11](https://github.com/ferryhe/excel_to_act/issues/11), noting source versions and unverified items. The latest third-party versions have not been reassessed, so do not claim their fidelity has been disproven.

### Items removed from this cycle

The dedicated docling_fidelity_assessment.md, comparison of three tool extensions, and ten-part source-evidence assessment are not conditions for this cycle's milestones. The original research acceptance was not carried out and must not be marked complete.

### Reopen criteria

Reopen when there is a clear need to integrate a document parser and a specific use case that current ingestion/view capabilities cannot support; then pin versions, inspect source, and validate fixtures against the dimensions required for that use case.

Related: [#3](https://github.com/ferryhe/excel_to_act/issues/3), project overview; [#11](https://github.com/ferryhe/excel_to_act/issues/11), current ingestion approach.

---

<!-- ISSUE_NUMBER: 11 -->
<!-- ISSUE: docs(design): align the existing pipeline, processing layers, and artifact boundaries -->

## Issue 11 docs(design): align the existing pipeline, processing layers, and artifact boundaries

[GitHub Issue #11](https://github.com/ferryhe/excel_to_act/issues/11) · Closed; PR #36 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P1; define the minimum boundary conventions required by #5/#13 first.

**Current delivery:** [PR #36](https://github.com/ferryhe/excel_to_act/pull/36) added the normative layer and artifact boundaries and aligned the related documentation without moving directories or adding runtime layers.

At the October 3 planning baseline, L0–L3 had no normative definitions. PR #36 delivered the boundary document without moving directories or adding wrapper layers.

### What to do

Create a short architecture document that defines the boundaries for fact ingestion, deterministic views, rule/semantic interpretation, and numerical verification, along with cross-layer shared services.

### Deliverables and scope

- docs/design/layered_architecture.md: responsibilities, inputs, outputs, and prohibited behavior for each layer.
- Use repository-relative paths in the file-level map and mark what is implemented versus planned.
- store, orchestrator, interfaces, and shared schemas are cross-layer services; report can render multiple phases and should not be forced into one processing layer.
- verify/completeness is a structural check at the decomposition boundary, distinct from the numerical verification in #7.
- List actual cross-phase Artifacts; do not introduce unused models merely to avoid a local metadata dictionary or achieve “no bare dicts.”
- Record the difference between rule classifications/confirmations already output by the orchestrator and the target semantic workflow in README Step 3; align classification enums and documentation goals with the actual contract.
- Distinguish pre-generation workbook-baseline verification from post-generation code-equivalence comparison.
- Record why document-parser integration is deferred under #10 and which evidence remains unverified.

Out of scope: directory refactoring, empty tools/agents/skills scaffolding, full semantic implementation, generators, and compression-algorithm research.

### How to verify

- [ ] Each existing module has a clear primary responsibility; dependency directions for cross-layer services are listed.
- [ ] The difference between the current call chain and the target workflow is clear; structural pass does not mean numerical pass.
- [ ] Inputs and outputs are defined for #5 object coverage, #13 factual fields, #14 views, and #7 numerical reports.
- [ ] The file map uses portable relative paths and accurately distinguishes existing and planned files.
- [ ] Differences between runtime classification/prompt enums and documentation are explained; enums are not added indiscriminately.
- [ ] Required rules match actual Artifacts, and differences are stated consistently in A1/README/PR plans.

This item only aligns the architecture; it does not replace runtime acceptance for implementation Issues.

---

<!-- ISSUE_NUMBER: 12 -->
<!-- ISSUE: docs(design): merge the Agent reading contract into #14 -->

## Issue 12 docs(design): merge the Agent reading contract into #14

[GitHub Issue #12](https://github.com/ferryhe/excel_to_act/issues/12) · Closed; requirements folded into #14 / PR #29

### Disposition

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`.

This item's requirements were merged into [#14](https://github.com/ferryhe/excel_to_act/issues/14), whose deliverables and acceptance criteria now include the following requirements; this item is no longer scheduled separately. It was closed because its work was merged, not because the Agent contract has been implemented.

### Complete requirements transferred to #14

- Define agent_reading_contract.md together with the view's view_id, stable record identifier, and source identity.
- The Agent locates real upstream artifacts from handoff and reads permitted deterministic views/projections.
- Output includes source_location; referenced views, source runs/versions, and records must all be valid.
- Report opaque items as unresolved; do not guess at sources or results.
- Factual values can be checked against their sources; clearly label deductions and conclusions that need verification. A schema cannot guarantee that reasoning is correct.
- Include at least three valid and three invalid output examples; missing sources, unknown view_id, wrong run/version, and invalid record references must fail.
- Reverse lookup for cells/ranges returns the complete sheet/address; workbook, VBA, and package parts use real object_id/ooxml_part values, never fabricated A1 references.

### Acceptance ownership

All of the above are accepted together under #14. This Issue retains its historical context and navigation entry; update #14 if rules need to change, so the two contracts do not drift independently.

---

<!-- ISSUE_NUMBER: 13 -->
<!-- ISSUE: fix(ingest): preserve raw values and formula properties and apply fidelity rules -->

## Issue 13 fix(ingest): preserve raw values and formula properties and apply fidelity rules

[GitHub Issue #13](https://github.com/ferryhe/excel_to_act/issues/13) · Closed; PR #28 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P0.

**Current delivery:** [PR #28](https://github.com/ferryhe/excel_to_act/pull/28) delivered field-level fidelity rules, schema export, and legacy `phase1.v1` read-back coverage. It does not add Decimal evaluation, Excel recalculation, or full dynamic-array evaluation.

At the October 3 planning baseline, raw value/type and date serial evidence had the documented gaps; PR #28 delivered the fidelity rules and schema/read-back work summarized above.

### What to do

Clarify the original documentation task as “field-level rules + minimal fidelity improvements to the current ingestion path.” Keep raw evidence alongside convenient parsed values to create an implementation task that can be accepted directly.

### Deliverables and scope

- docs/design/fidelity_rules.md: actual fields, permitted conversions, counterexamples, and checks for each rule, distinguishing satisfied rules from gaps.
- Add the necessary fields to the existing CellInventory/manifest/SourceLocation: raw value text, raw cache text, raw OOXML type, date system, and raw type, properties, and range for shared/array formulas. Fix exact field names and compatibility behavior in the rules.
- Distinguish original text from expanded shared-formula results; array-formula text is currently retained, but do not claim that ref/spill information is fully saved.
- Preserve the meaning of error values, percentage number_format, and cache availability; a displayed date does not replace the source number.
- Align JSON Schema with existing store read-back and add a minimal fidelity regression fixture.

Out of scope: a Decimal evaluator, Excel recalculation, Python generation, format rendering, and full dynamic-array evaluation. Do not reimplement the ordinary cache ingestion completed in #4.

### How to verify

- [ ] Every rule has an actual field and a runnable assertion; a rule cannot be marked satisfied without a field to carry it.
- [ ] Dates retain their source serial, source type, and date system; the displayed value does not overwrite the raw value.
- [ ] Original numeric text and number_format are traceable; the value is not reduced to a float alone.
- [ ] Original formula text, shared/array types and ranges, and permitted expansion conversions are traceable.
- [ ] Raw cached text is distinguished from parsed values; 0/False/empty string are not mistaken for missing values.
- [ ] New fields are exported to Schema, and old artifact read-back behavior is explained and verified.
- [ ] Consistent with existing ordinary cache ingestion and coverage accounting; tests and Ruff pass.

Dependencies: the required source conventions from #11; align with #5 by object responsibility. Completing this Issue provides the factual basis for reliable views in #14 and value-type comparison in #7.

---

<!-- ISSUE_NUMBER: 14 -->
<!-- ISSUE: feat(views): deterministic slice views and verifiable Agent references -->

## Issue 14 feat(views): deterministic slice views and verifiable Agent references

[GitHub Issue #14](https://github.com/ferryhe/excel_to_act/issues/14) · Closed; PR #29 merged

### Current status and priority

Review baseline: 2026-10-03, branch `feat/step1-coverage-handoff`, commit `c86fa55`. Priority: P1.

**Current delivery:** [PR #29](https://github.com/ferryhe/excel_to_act/pull/29) delivered deterministic views, source/reference validation, and the design and Agent reading contracts. This includes the requirements folded from #12; #12 had no separate implementation.

At the October 3 planning baseline, the view contract and compiler were not implemented. PR #29 delivered the scoped view path and included the reading/write-back contract folded from [#12](https://github.com/ferryhe/excel_to_act/issues/12).

### What to do

Generate deterministic views from the inventory/graph referenced by handoff, retaining sources, supporting reverse lookup, and enforcing a budget boundary; validate references in Agent output.

### Deliverables and scope

- docs/design/views_l1_compiler.md and agent_reading_contract.md: jointly define the interface, view_id, stable record identifier, source-artifact identity, and SourceLocation.
- Reuse the view contract for Artifact/CellInventory/SourceLocation and a minimal executable compilation path that slices by sheet/region.
- Locate inputs from handoff and verify the sources and identities of the upstream artifacts actually used; Agents read only permitted deterministic projections and do not guess about opaque items.
- The first release uses ordered slices and budget splitting; specify whether budgets are estimated or precisely counted with a fixed tokenizer, and define behavior for a single over-budget item without silently truncating it.
- Use the real graph from #6 for dependency regions; unresolved references remain marked unresolved.
- Validate Agent output: view/source run or version/record references exist, factual values match their sources, and deductions and conclusions needing verification are clearly labeled.
- Include at least three valid and three invalid output examples, along with runnable determinism and source-lookup tests.

Defer deduplication of identical formulas/formats until needed; do not implement relative-formula normalization or a full compressor, wait for separate A3 research, or include model calls, prompt tuning, or Python generation.

### How to verify

- [ ] Two compilations with the same input and configuration produce identical bytes.
- [ ] Each cell/range record can be traced back to the full sheet/address; workbook, VBA, and package parts trace to real object_id/ooxml_part values, with no fabricated A1.
- [ ] All records remain traceable after budget splitting; over-budget cases neither lose records nor fabricate sources.
- [ ] At least one cross-sheet dependency sample has no false local edge or incorrect reference.
- [ ] Runnable validation covers at least three valid and three invalid Agent outputs.
- [ ] Missing source_location, unknown view_id, wrong run/version, and invalid record references all fail; opaque items are reported explicitly.
- [ ] Reuse existing Artifact, align Schema and upstream source identity, and pass tests and Ruff.

Dependencies: #11 minimal inter-layer conventions, the trusted-fact contracts in #5/#13, and reference-graph acceptance in #6. A single-sheet slicing prototype can proceed earlier, but reliable dependency slicing cannot be declared complete until dependencies are met. Numerical verification in #7 is independent of this item.

---

## Local implementation records

The entries below use descriptive `LOCAL-*` labels; historical #15–#18 were local implementation-record numbers, not GitHub Issue IDs. GitHub #15–#17 are PRs, including [PR #15](https://github.com/ferryhe/excel_to_act/pull/15) and [PR #16](https://github.com/ferryhe/excel_to_act/pull/16); GitHub #18 is a separate real Step 2 Issue outside this Epic. Preserve the boundary between completed and incomplete work, and do not create duplicate Issues.

### LOCAL-DATA-TABLE What-If Data Tables

The current branch has ingest/data_table.py: it reads `<f t="dataTable">` directly from the top-left cell, retains text, ref, r1/r2, row/column input direction, and one-/two-variable information, and writes and counts it in the sheet as RangeInventory(kind=data_table). The latest commit, c86fa55, fixes interpretation of one-variable row/column inputs. Tests cover no table, one-variable row/column, two-variable, and original openpyxl behavior.

What-If Data Table evaluation is not included. Unified fidelity for raw formula properties and array ranges belongs in #13; object coverage accounting belongs in #5. Implementation evidence: ingest/data_table.py, inventory/extractor.py, and tests/test_data_table.py in PR #15.

### LOCAL-FORM-CONTROLS VML control bindings

The current branch has ingest/form_controls.py, which reads FmlaLink, FmlaRange, and FmlaMacro from vmlDrawing, excludes Note/Pict, and writes form_control ranges and metadata; package relationship target parsing and VML samples are covered by tests.

Boundary: independent parsing/fallback for ctrlProps and dedicated input classification for linkedCell do not meet all originally intended acceptance criteria; ActiveX remains opaque. This records only VML binding capability and does not claim that “all form-control acceptance criteria are complete.” Track whether to extend these capabilities based on actual workbook needs. Implementation evidence: ingest/form_controls.py and tests/test_form_controls.py in PR #15.

### LOCAL-VBA VBA source and candidate references

The current branch has an optional oletools extractor, VbaModule, candidate references, and vba_ref edges in FormulaGraph, all using the same node namespace. References that cannot be determined statically retain a low-confidence/unresolved boundary; a warning is recorded when the dependency is not installed. Existing tests include candidate references and the fallback path.

An end-to-end sample using a real xlsm containing vbaProject.bin is still needed; mock tests do not count as acceptance of real extraction. No real `.xlsm` VBA integration path was demonstrated. VBA is not executed, and this does not cover all macro semantics or XLM/DDE. Implementation evidence: ingest/vba.py, inventory/vba_links.py, orchestrator/phase1.py, and tests/test_vba.py in PR #15.

### LOCAL-HANDOFF independent checks and handoff artifacts

The current branch has verify/completeness.py, CompletenessReport, Handoff, report/handoff.py, store writes, and a CLI failure exit path. It outputs completeness.json, handoff.json, and a short English handoff.md; the run directory and root alias can be read back. Existing tests cover basic artifacts, missing sheets, and CLI failure.

Structural checks are not numerical verification. `CoverageSummary.discovered_workbook_objects` comes from the independent source scan, and recognized count is based on matched source identities. The summary counts alone do not prove identity-by-identity coverage; #5's `verify_completeness` check compares identities and blocks omissions. #8 covers failed runs for input files that fail during reading. Implementation evidence: PR #15 and the related files in verify, report, store, orchestrator, and CLI.
