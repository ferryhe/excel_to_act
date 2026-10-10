## What to deliver

Deliver `excel-to-act step2 prepare`: package validated Step 1 / Step 2 artifacts into a directory that an agent can read incrementally, and generate English human-readable and machine-readable handoffs. Later queries reuse this prepared result instead of reparsing the original Excel file for each question.

This is the first of three sequential implementation items and is recommended as one PR. Existing #6, #13, #14, and #18–#21 are foundations and should not be reimplemented. Baseline `main` is `9251019a3b77fe12ce36a811d7392bc781219d0d`; the current local working tree also contains reviewed but unpublished Step 1 namespace fixes and checkbox/ActiveX/VBA handoff changes. Before implementation, confirm these prerequisites are available and identify prerequisite commits or PRs in the PR; do not redevelop the same parsers.

## CLI contract

```text
excel-to-act step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]
```

- Inputs are the native `index.json` and corresponding Step 1 root; existing index commands continue to handle native `step1.v1` / `step1.batch.v1`.
- `--scope` is a user scope decision in `analysis.scope.v1` bound to the source SHA/run, not an AI handoff or data snapshot. If omitted, retain all worksheets.
- Check source identity, index references, and consistency between the scope worksheet list and its decision. For batches, retain all entries and diagnostics for failed entries; do not silently drop them.
- `--dry-run` returns input, capability, and plan files without creating directories/files or restoring history.
- `--resume` decides reuse using source references, scope, compiler version, parameters, and artifact hashes. Ordinary preparation and reading do not consume the existing three-attempt indexing recovery budget.
- stdout returns JSON containing at least status, source/run identity, artifact references, diagnostics, and next action. Integrity failures return structured diagnostics and a nonzero exit code.

## Deliverables

1. A runnable and discoverable prepare CLI; `step2 tools` `reader_commands` lists only capabilities implemented by this PR.
2. A reading package:

   ```text
   reading/
     manifest.json
     scope.json
     dependency_audit.json
     retained_dependency_cells.json
     step1/HANDOFF.json
     step1/HANDOFF.md
     step2/HANDOFF.json
     step2/HANDOFF.md
     sources/<source-id>/views/<view-id>.json
     sources/<source-id>/dependency_graph.json
   ```

3. The manifest binds source/run, native index, scope, compiler version/parameters, view and graph hashes, plus lookup information and allowed dependency scope for later selectors. References use explicit roots and relative paths; preserve `SourceLocation` exactly as recorded.
4. Reuse existing `WorkbookView` / `ViewRecord` / `SourceLocation`. During preparation, include checkbox/ActiveX artifacts in canonical sheet views only after native index hash, typed model, and workbook identity checks. A complete typed record is a fact; determine location from actual owner/part/shape fields, and do not fabricate an A1 location for a linked cell as though it were the control itself. Added reading records do not change the Step 1 coverage denominator.
5. Generate a dependency audit/snapshot from validated inventory, name declarations, and supported static references, reusing source-data reads performed during preparation. Bind output to source SHA, run, and scope hash; register paths and SHA-256 in the manifest. Sidecar filenames without hashes in an older scope are historical hints only, not authoritative inputs; do not import manually edited dependency values.
6. Render English human-readable and machine-readable handoffs from one validated context. Distinguish source quality, reference integrity, static readiness, and unverified runtime/numeric behavior. A changed scope creates a new reading revision; do not edit checksum-bound final artifacts.
7. Add necessary runtime models / JSON Schema, README and agent preparation-stage instructions, and CLI regression tests. Include the reviewed `docs/design/progressive_excel_exploration.md` in the first PR and implement this Issue's CLI contract in the documentation.

## Scope and dependencies

- This item ends at the preparation package and handoff: do not add query/trace, agent-claim validation for evidence packets, business interpretation, macro execution, recalculation backends, or Python generation.
- Do not replace native index/source handoff; do not change Step 1 fidelity, opaque, or promotion semantics; do not lower existing checks when preparation fails.
- Retain references to original source/inventory and VBA files instead of copying large source artifacts. Save views as separately readable files and reuse the graph; do not create a new database or agent framework for future questions.
- Report missing caches and dynamic references explicitly. A static audit must not claim all dependencies are known, and a cache must not be described as the result for changed inputs.
- Item 2 consumes this item's manifest/canonical views; item 3 consumes this item's saved graph. This item provides the required data contract but does not implement later query/traversal algorithms early.

## Checks

- [ ] CLI help, tool catalog, returned JSON, and failure exit codes match the contract.
- [ ] Incorrect source/run scope, conflicting worksheet lists, and missing/changed native references produce structured failures; original finals and native index bytes remain unchanged.
- [ ] Dry-run performs no writes with either an absent or existing output directory; recovery history is not changed.
- [ ] An unchanged resume reuses results; changes to scope/source/compiler/options invalidate them correctly; changed or missing preparation artifacts are not reused.
- [ ] The reading package can be relocated using explicit roots; preserve original source locations without relying only on absolute paths from the current machine.
- [ ] Checkbox/ActiveX canonical records can be verified against original module artifacts and retain complete facts/source identity; bindings do not enter the logical-object coverage denominator.
- [ ] Editing an unreferenced old sidecar cannot change generated data; editing or removing a generated snapshot/audit causes reuse/validation to fail.
- [ ] For the same Pricing source version and four approved excluded tabs, generate 864 supported static inbound references, 12 ranges, and 1,464 stored-cell records (including 1,364 formula cells) from validated source data; preserve blank-cell and cache-availability semantics without guessing values or recalculating.
- [ ] Fixture CI and local real Pricing CLI runs both save inspectable evidence; real inputs and large outputs remain in the local artifact directory and are not committed.
- [ ] Record actual inventory-parse and graph-build counts to prove preparation artifacts can be reused; related tests, Ruff, and CI pass.
