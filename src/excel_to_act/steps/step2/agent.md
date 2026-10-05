# Step 2 host-agent contract

## Normal flow

Run `excel-to-act step2 index --handoff PATH --step1-root DIR [--out DIR] [--resume]`. The command resolves the current `step1.v1` source handoff or `step1.batch.v1` batch handoff, builds `index.json` and `INDEX.md`, then validates the saved index and its referenced artifacts. The default output directory is `output/step2_index/`. It retains every batch entry, including failed entries and entries without a readable per-source handoff; never filter or drop an entry because it failed or is incomplete.

`step2 index` returns `status`, `diagnostics`, `metrics`, `retryable`, and `next_tool`. Its `status` reports source/batch status unless validation blocks the result. A usable partial index exits successfully; `blocked` exits nonzero. Standalone `step2 validate` returns validation integrity in its response `status`. For the persisted view, read `index.json`: `validation_status` is integrity status, top-level `status` is batch status, and each `entries` item has its source `status` and `next_actions`. Diagnostics identify the source entry and artifact when available. `retryable` and `next_tool` recommend only existing actions; per-entry `next_actions` may name a Step 1 action for the host to invoke through Step 1's own recovery contract.

## Tools and recovery

Use `excel-to-act step2 tools` to read the action catalogue. `excel-to-act step2 tool NAME ...` dispatches exactly one listed action and returns its result; the host owns the loop. The normal `step2 index` command already performs resolve → build → validate. Individual catalogue actions are `step2.handoff.resolve` → `step2.index.build` → `step2.index.validate`.

Use `--resume` with `step2 index` or `step2.index.build` to reuse validated entries whose input handoff and referenced artifacts have not changed. Keep calls using one output directory sequential. Step 2 records a maximum of three actual catalogue dispatches in the shared `state.json` attempt history. Stop before a fourth dispatch, on a repeated tool/input pair, after two consecutive attempts without progress, or after a non-retryable failure. An interrupted build can be resumed within the remaining budget. After the bound is exhausted, retain its evidence and use a new output directory for a new session. Do not count Step 1 recovery calls against Step 2's attempt history; follow Step 1's `attempt_history.json` and recovery limit for those calls. Never ask Step 2 to invoke Step 1 tools.

## Downstream analysis handoff

For downstream analysis, treat `index.json` and `INDEX.md` as navigation and quality context. Use each entry's stable source identity, status, metrics, diagnostics, artifact references, and next actions to select relevant Step 1 artifacts. Load detailed artifacts progressively, only when needed for analysis; do not copy every artifact into the index or rescan the whole batch. Preserve all entries when selecting work. The analysis agent chooses its analysis tools and loops on their results. Keep analytical findings and conclusions in its own analysis output; Step 2 records indexing, reference integrity, and source quality context only.

`output/step2_index/index.json` is the new batch/source navigation index. It is distinct from the legacy per-workbook `artifact_index.json` produced by `store/local_store.py`.
