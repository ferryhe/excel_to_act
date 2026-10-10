## What to deliver

Deliver `excel-to-act step2 query` and extend `excel-to-act views validate`: an agent reads Excel incrementally through small source-evidence packets, then validates its citations. It should not repeatedly load the large views file for the entire workbook.

This is the second of three sequential implementation items and is recommended as one PR. It depends on the manifest, canonical views, scope, and allowed dependency scope from [#30 prepare](https://github.com/ferryhe/excel_to_act/issues/30); do not reimplement preparation or extraction.

## CLI contract

```text
excel-to-act step2 query --manifest MANIFEST --source-id SOURCE --kind KIND [--target TARGET] [--sheet SHEET] [--range A1_RANGE] [--budget N] [--cursor CURSOR] [--out PACKET]
excel-to-act views validate --views EVIDENCE_PACKET --output CLAIMS
```

- `--source-id` is always required, including for a single-source reading package. Get it from the manifest source list; do not automatically select the first source.
- Initial kinds: `overview`, `sheet`, `cell`, `range`, `name`, `control`, `vba`, and `feature`.
- On read, verify source identity, bound manifest, and selected evidence hashes. Load only the required view files; do not reparse raw Excel or the entire single-file views artifact.
- `overview` / `sheet` return mechanical summaries and bounded pages; `cell` / `range` preserve value, formula, cache, availability, and complete source location; `name` preserves declarations and scope, and returns candidates plus `needs_selection` for ambiguity instead of guessing a definition.
- `control` returns complete facts from the canonical records in item 1; initial `vba` support returns the complete module record and source-file reference, not a snippet presented as a complete source fact.
- Use the existing budget estimation method and do not split records. Pagination explicitly reports delivered IDs, next cursor, omissions, and oversized status; for oversized records, return a reference and next step instead of silently truncating.
- Distinguish no match, ambiguity, unsupported, excluded scope, truncation, missing data, and integrity failure. Blocking/validation failures return structured diagnostics and nonzero exit codes; reads do not consume indexing recovery budget.

## Deliverables

1. A runnable query CLI and eight selector kinds; the tool catalog lists reader capabilities implemented by this PR, their parameters, results, and next actions.
2. Source-bound evidence packets containing canonical view references, actually delivered records with original IDs/facts/`SourceLocation`, selection metadata, pagination/budget details, and diagnostics. Mark mechanical summaries as derivations; do not fabricate source facts.
3. Scope checks at the selector layer:
   - Normal facts from retained sheets are readable; `overview` / `feature` for excluded sheets show only availability, exclusion status, counts, and allowed scope.
   - Cells on excluded sheets must be within the approved dependency scope; a requested range must be wholly contained in the approved scope, and mere overlap is insufficient.
   - An out-of-scope request returns `out_of_scope`, the requested selector, and the allowed narrower target; do not silently clip it.
   - Name destinations obey the same restriction. Excluded controls and worksheet VBA provide metadata only, not interaction/code content for downstream analysis; return `needs_scope_resolution` when ownership is unknown.
   - Expanding the read boundary requires a newly confirmed scope and prepare revision; do not add a default switch that bypasses scope.
4. The views validator supports both the old list-of-views format and new packets. After validating canonical reference hashes and selected records, reuse existing `validate_agent_output`: only actually delivered records may be cited; facts must be complete objects; reject wrong run/location, tampered facts, and records not provided. PASS means source references are consistent, not that the semantics or numbers are correct.
5. Update the Step 2 agent / README with the sequence: manifest identity → overview → specific question → targeted query → citation validation; do not repeat a read when no new evidence is needed. Documentation for this PR must not present trace as available.
6. CLI regressions and source-workbook/scope/control/VBA fixtures; maintain compatibility with existing view-reference validation.

## Scope and dependencies

- This item handles selection, packets, read scope, and citation validation; it does not recursively trace upstream/downstream, expand dependencies for name/range members, or build a complete VBA call graph.
- Reuse the data and hash contract from item 1 prepare; do not change extraction, native index, promotion, or source accounting.
- Interpreting VBA workflows, TypeSafe judgments, and runtime proof belong to later analysis; this item only provides accurate evidence.
- Do not add a token-compression system, database, search service, graph UI, or full Excel calculation engine.

## Checks

- [ ] Omitted/invalid source IDs have controlled behavior; names shared across sources or scoped to a workbook/worksheet require sufficient context and never select the first match.
- [ ] `overview` / `sheet` / `cell` / `range` / `name` / `control` / `vba` / `feature` each has a CLI success case and corresponding real failure case.
- [ ] A missing stored cell, unavailable cache, and zero value are distinct; do not fill with zero or change the fact object/source location.
- [ ] Pagination can continue without omitting or duplicating selected records; budget/oversized information is explicit, and a complete record is never truncated into a false fact.
- [ ] `CalculationOfBE_CI!O7:Q112` is readable within the approved scope; `CalculationOfBE_CI!A1` and broad queries extending beyond the approved scope return `out_of_scope`; excluded-sheet overview does not return ordinary cell/formula pages.
- [ ] On Pricing, `CalculationOfBE_OtherDisease` `Check Box 1` returns canonical binding `CalculationOfBE_OtherDisease!F14`; its linked cell can be read separately. `PremiumTable.CmdPremium` and the complete `Sheet9` source record are readable.
- [ ] Real source-backed binding/cell/module references pass; changing target/facts, wrong run/location, corrupt canonical reference, or citing an undelivered record all fail with nonzero exit status.
- [ ] The old `views validate` list format still works; derivation / unverified / opaque data is not misclassified as runtime or semantic PASS.
- [ ] Measurements across multiple small questions show no repeated raw extraction, full inventory parse, or graph build, and no repeated load of the approximately 773 MB single-file Pricing views artifact; record selected views and actual bytes read.
- [ ] Ordinary query, pagination, and validation do not exhaust the existing three-attempt indexing recovery history; related tests, Ruff, and CI pass.
