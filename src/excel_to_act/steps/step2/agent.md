# Step 2 host-agent contract

Author new instructions, documentation, report narration, and generated artifact prose in English. Preserve source values, formulas, identifiers, and immutable provenance verbatim. Only the repository-root `README.md` and `README.zh-CN.md` are approved bilingual documentation exceptions.

## Normal flow

Run `excel-to-act step2 index --handoff PATH --step1-root DIR [--out DIR] [--resume]`. The command resolves the current `step1.v1` source handoff or `step1.batch.v1` batch handoff, builds `index.json` and `INDEX.md`, then validates the saved index and its referenced artifacts. The default output directory is `output/step2_index/`. It retains every batch entry, including failed entries and entries without a readable per-source handoff; never filter or drop an entry because it failed or is incomplete.

After indexing, run `excel-to-act step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]`. Prepare accepts only the native saved index, validates its references through the read-only integrity verifier, and does not dispatch or change the three-attempt recovery history. A supplied scope must be an `analysis.scope.v1` decision bound to exactly one indexed source SHA and run. Without a scope, all sheets are retained. Prepare preserves failed batch records and diagnostics in both handoffs; use the manifest's root-relative references and allowed-read paths when consuming the package.

`--dry-run` validates the inputs and reports planned package files without writing. Its metrics count inventory model deserializations performed by validation; raw workbook parses and graph builds remain zero. `--resume` reuses an unchanged package only after checking the scope, source/index references, compiler/options, the complete required output set, and every recorded output hash. A tampered, missing, or unlisted required prepared output blocks reuse. Static dependency output contains saved source values and cached-value availability only; it does not claim dynamic dependencies are complete, execute macros, recalculate formulas, or establish changed-scenario behavior.

The Step 2 handoff's `artifacts` entries use root-relative references. Audit and snapshot references include SHA-256 and byte length; the manifest reference contains only its explicit `reading` root and path because a manifest cannot checksum itself. Source rows keep each source's status, readiness, and diagnostics, including failed batch entries.

## Progressive evidence reading

Use the selected source identity from `manifest.json`, then ask for `overview` before reading a specific sheet or fact. Every query requires the exact `--source-id`; do not infer the first source in a batch.

```text
excel-to-act step2 query --manifest output/reading/manifest.json --source-id SOURCE --kind overview
excel-to-act step2 query --manifest output/reading/manifest.json --source-id SOURCE --kind cell --sheet PremiumTable --target B7
excel-to-act step2 trace --manifest output/reading/manifest.json --source-id SOURCE --kind name --target GP --direction upstream --out trace-packet.json
excel-to-act views validate --views trace-packet.json --output claims.json
```

Use a targeted `sheet`, `cell`, `range`, `name`, `control`, `vba`, or `feature` query for the question at hand. Keep the returned packet with the claims and validate citations against that packet. A packet validates only that cited facts, IDs, locations, run identity, and canonical files match delivered evidence; it does not validate semantic or numerical conclusions. A truncated query page returns a cursor. Continue it, increasing `--budget` when the next complete record is too large; never treat an omitted record as evidence. Respect excluded sheets and approved dependency ranges; a broader read needs a new confirmed scope and preparation revision.

Use `step2 trace` after a small query to follow static dependencies of a cell, range, or name. `upstream` follows existing formula-cell → referenced-input edges; `downstream` finds consumers, including names and containing stored ranges; `both` combines them. Name destinations and stored range members are labeled `structural_derivation` with manifest/file/record provenance, never source formula edges. Names require one declaration within the selected scope. Constants terminate; relative/dynamic/external/multi-area names remain unresolved. Current-row table selectors without caller context also remain unresolved.

Default bounds are depth 8, nodes 100, edges 200 (maximum 100/10000/20000). A trace returns frontier and next actions; continue from one frontier target with explicit bounds. It uses visited sets and loads selected prepared views once per call. Repeated reads verify prepared files again and never reparse the workbook, rebuild the formula graph, or consume indexing recovery attempts. Excluded-sheet formulas are never expanded; only approved read-only dependency cell records may be delivered, unchanged even when the record contains formula text. Approved name destinations, stored range members, and incoming formula edges from retained sheets remain traversable within the same bounds. VBA literal references remain auxiliary, with no proven read/write direction, execution order, or full call graph; `Range(Var1)` is unresolved. Inspect VBA separately with `query`.

Validate the trace packet before Step 3 interpretation. The validator replays its bounded structural derivation against canonical evidence and uses the existing source-record citation checks. `closure: not_proven` and `runtime_verified: false` remain explicit: cached values are saved results, never new scenario recalculation. Exported modules and event associations are declarations, never macro execution evidence. Give the exploration specification, packets, commands, and check results to a fresh independent reviewer. Accept only reproducible findings tied to this CLI/data contract. Investigate concrete unresolved evidence; a bounded TypeSafe/Jev latest judgment can adjudicate disagreement but cannot become source, numeric, runtime, or CI PASS evidence. See `docs/design/step2_trace_exploration.md` for replay and reviewer instructions.

`step2 index` returns `status`, `diagnostics`, `metrics`, `retryable`, and `next_tool`. Its `status` is index usability (`partial` for any non-empty built index) unless validation blocks it; an empty or blocked index returns `blocked`. Standalone `step2 validate` returns validation integrity in its response `status`. In saved `index.json`, top-level `status` is index usability, `validation_status` is validation integrity, and `entries[*].status` is each Step 1 source status; use each entry’s `next_actions` for its next actions. Diagnostics identify the source entry and artifact when available. `retryable` and `next_tool` recommend only existing actions; per-entry `next_actions` may name a Step 1 action for the host to invoke through Step 1's own recovery contract.

## Tools and recovery

Use `excel-to-act step2 tools` to read the action catalogue. `excel-to-act step2 tool NAME ...` dispatches exactly one listed action and returns its result; the host owns the loop. The normal `step2 index` command already performs resolve → build → validate. Individual catalogue actions are `step2.handoff.resolve` → `step2.index.build` → `step2.index.validate`.

Use `--resume` with `step2 index` or `step2.index.build` to reuse validated entries whose input handoff and referenced artifacts have not changed. Keep calls using one output directory sequential. Step 2 records a maximum of three actual catalogue dispatches in the shared `state.json` attempt history. Stop before a fourth dispatch, on a repeated tool/input pair, after two consecutive attempts without progress, or after a non-retryable failure. An interrupted build can be resumed within the remaining budget. After the bound is exhausted, retain its evidence and use a new output directory for a new session. Do not count Step 1 recovery calls against Step 2's attempt history; follow Step 1's `attempt_history.json` and recovery limit for those calls. Never ask Step 2 to invoke Step 1 tools.

## Downstream analysis handoff

For downstream analysis, treat `index.json` and `INDEX.md` as navigation and quality context. Use each entry's stable source identity, status, metrics, diagnostics, artifact references, and next actions to select relevant Step 1 artifacts. Load detailed artifacts progressively, only when needed for analysis; do not copy every artifact into the index or rescan the whole batch. Preserve all entries when selecting work. The analysis agent chooses its analysis tools and loops on their results. Keep analytical findings and conclusions in its own analysis output; Step 2 records indexing, reference integrity, and source quality context only.

`output/step2_index/index.json` is the new batch/source navigation index. It is distinct from the legacy per-workbook `artifact_index.json` produced by `store/local_store.py`.

## Reviewed conversion checkpoint

Only after Stage 1 has both current approvals, create the read-only Stage 2 checkpoint with `excel-to-act step2 report --index INDEX --step1-root STEP1_ROOT --workflow WORKFLOW [--source-id ID]`. It validates the existing saved index and selected source against the approved Step 1 run; it does not rebuild the index or modify Step 1/2 artifacts. Review the JSON/Markdown report and its source/index hashes before recording the Agent decision. Record a human decision only after an explicit user response. Stage 3 cannot proceed until both Stage 2 decisions approve the current report pair. Use `excel-to-act workflow status --workflow WORKFLOW` to inspect freshness and stage gates.

Record Stage 2 decisions with `excel-to-act workflow confirm --workflow WORKFLOW --stage 2 --reviewer agent|human --decision approve --message TEXT`. The human decision requires an explicit user response. Route a rejection with `excel-to-act workflow reject --workflow WORKFLOW --stage 2 --reviewer agent|human --return-to 1|2 --message TEXT`; return to Stage 1 if the accepted source checkpoint is wrong. Check status after any decision. Revisions invalidate downstream approvals.

## Shared workflow and decision order

Treat `input/` as the source Excel area and `output/` as the home for each new source-bound workflow and its reports. Start a new conversion from the current workbook and user request; never carry case facts or approvals from another workbook or run.

Read `stepN tools` and choose the existing actions that answer this case's open questions. Explore from the source and current bound evidence, inspect each action's checks and outputs, then prepare the paired machine-readable JSON and reader-facing Markdown handoff. Review all files named in the handoff and confirm their source and upstream bindings before asking for a decision. Do not continue when a required check fails or an input is stale.

Before entering a stage, run `workflow status` and verify that the stage is currently permitted and every required upstream decision and artifact is current. A successful tool check, report creation, draft CLI action, or clarification does not itself approve a stage or advance the workflow. Keep legacy standalone source-exploration APIs available for discovery; they do not create a workflow approval.

The default review order is Agent first, then the actual human. `workflow confirm` records approvals only; use the stage-specific `workflow.reject` command and its explicit `--return-to` route to reject. Record the Agent decision against the exact current JSON/Markdown pair with `workflow confirm`; present that pair and its practical limits to the user; record a human approval only after an explicit response. New TypeSafe approvals also require a current matching Agent approval. Use TypeSafe only under a valid, explicitly scoped authorization already registered for this workflow; never infer or transfer delegation. The Step 3 input-boundary checkpoint always requires an actual human decision.

Ask the user when target, scope, input kind, axis, units, or business interpretation is materially unresolved. Keep unknown facts marked as unknown; do not turn planning into measured execution or claim checks that did not run.

When a reviewer rejects a report at Stage 2, route it to the current or an earlier responsible stage with `workflow reject --workflow DIR --stage 2 --return-to 1|2 --message TEXT`. Keep the rejected artifact and decision in the ledger. Revise affected reports, rerun their checks, refresh later reports that depend on them, and obtain fresh Agent and human or explicitly delegated reviews before advancing. Inspect the ledger with `workflow status` after each decision and before each transition.
