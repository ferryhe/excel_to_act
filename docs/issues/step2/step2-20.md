**CLI:** `excel-to-act step2 tool NAME ...`; `excel-to-act step2 index ... --resume`

## Goal

Allow the host agent to invoke one selected tool and resume an interrupted or partially completed batch without reprocessing unchanged entries.

## Scope

- Add `step2 tool NAME` dispatch for tools exposed by `step2 tools`.
- Persist an execution state with input handoff hashes, per-entry status, output hashes, and an attempt history for Step 2 catalogue actions.
- Add `--resume` to skip entries whose inputs and referenced artifacts are unchanged.
- Make `step2 tool NAME` dispatch exactly one selected Step 2 action and return its result; the host agent owns the loop, follows Step 2 `next_tool` or Step 1 `next_actions`, and revalidates only the affected entry.
- Stop Step 2 mutating repairs after three attempts, two consecutive attempts without progress, a repeated tool/input pair, or a non-retryable failure.
- Keep Step 1 repair calls outside Step 2's `state.json` attempt count. When the host agent follows a Step 1 `next_actions` entry, it must use Step 1's own persisted `attempt_history.json` and recovery limit.
- Preserve partial indexes and failure evidence so a later corrected run can continue.

## Files

- `src/excel_to_act/steps/step2/workflow.py`
- `src/excel_to_act/interfaces/cli.py`
- `src/excel_to_act/steps/step2/README.md`
- `tests/test_step2_index.py` and `tests/test_step2_cli.py`

## Acceptance criteria

1. Resuming an unchanged batch skips already validated entries.
2. Changing one source entry causes only that entry to be reprocessed.
3. Manual and host-agent-orchestrated Step 2 tool calls share one persisted attempt limit.
4. Step 2 attempt history counts Step 2 catalogue actions only; Step 1 action history stays under Step 1's recovery contract.
5. `step2 tool` performs one dispatch and does not automatically execute a multi-tool loop.
6. The loop stops on its configured bounds and returns the evidence and stop reason.
7. Opaque content does not trigger repeated parser calls; it remains visible with a human or specialist next action.

## Dependencies and boundary

- Depends on #18's stable per-entry identity and #19's validator result.
- The host agent chooses among listed tools. The Python CLI runs deterministic tools and does not invoke a language model.
