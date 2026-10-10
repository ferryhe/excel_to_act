# Fixed-scenario GP calculation

Compute `GP` at `Premium!J1` independently from the original workbook's inputs and required formulas, then reconcile it with fresh native Excel results. Use the saved scenario: age 0, Male, benefit term 30, premium term 10, sum assured 1000, pricing interest 0.035 and CI decrement Yes. Preserve the original workbook and prior source-bound artifacts.

## Pre-implementation assessment

Tier: **complex**, persistent worker `gpt-6-sol / high`. GP requires coupled annual states, dynamic names and lookup columns, computed CI rate tables, the MultipleCI Markov recurrence, and a source-bound runtime trace. An explicit Premium-only pipeline would still duplicate coupled upstream formula families or incorrectly consume formula caches. A narrow lazy evaluator can compute the selected return cells and branches, record actual dependencies and preserve the formulas. This is an execution prototype, not general Python generation or a complete Excel engine.

The architect confirmed that the current selected benefit columns are `Level100` and `None`; GP-dependent `ROP` and `WaiverOfPrem` columns are not selected. An active cycle must block, rather than borrowing a saved cache or silently inventing an iteration solver. The source workbook's native iterative settings remain recorded for the independent oracle.

Four previously optional-excluded sheets contain required `*CITableFinal` lookup bridges at `O7:Q112`. The new target calculation reads only active bridge cells and their required prerequisites. The old scope file and Step 1/2/3 source bindings remain unchanged. This new target scope is recorded explicitly.

## Smallest complete tool flow

1. Capture a native Excel baseline in an isolated read-only copy, with events and macros disabled, using `CalculateFullRebuild`. Record source hash, engine/settings, targets and requested annual ranges. Never save over the input workbook.
2. Calculate the requested target using original formulas and non-formula values. Reuse `openpyxl.Tokenizer` for lexical tokens; unsupported active functions, syntax, external dependencies and cycles remain explicit blockers. Formula caches are never an evaluation input.
3. Write a result and an active-dependency/binding report. Show raw and calculated cells, selected names/ranges/lookup columns, branch decisions, source error handling and evaluation order. Explain GP modules, arrays and sequential recurrences from this report.
4. Reconcile GP, its three root intermediates, annual state/timing vectors and five cost vectors against the independent baseline. Compare source/scenario identity before values, declare numeric tolerances and retain exact error mismatches.
5. Publish the actual reusable command syntax in the CLI catalogue and README pair. Keep numeric execution separate from the existing static Step 3 validator; citation validation is not runtime equivalence.

The numerical command accepts a workbook, an exact target and an output directory. The oracle command accepts a workbook, target names and finite ranges. A different Excel can use the same supported flow; unsupported active formula features must fail visibly. No new plugin system, general solver, GPU layer or all-configuration promise is required.

## Acceptance

- Independent Python GP matches the fresh Excel fixed-scenario baseline, currently `2.7283284514524744`; the baseline is comparison data only.
- AnnuityDue `8.574676381865864`, PVLoading `3.161998467798486`, and PVFB `14.767563151498457` also reconcile, with annual arrays checked rather than only the final ratio.
- Actual dynamic ranges are `T9:T18` and `CD10:CD19`; loading uses the selected payment-term column and its original approximate lookup behavior.
- Preserve the source's unusual `I11` reference to the prior H state and its `I_t * I_t` CI cost. Do not repair actuarial formulas without separate authorization.
- Test meaningful synthetic lookup, branch, error, recurrence, unsupported-feature, source-binding and active-cycle cases. Unsupported implementation capabilities cannot be swallowed by source IFERROR.
- CLI output and saved JSON/Markdown agree; paths and source hashes are validated, originals remain unchanged, and all authored prose is English.
- Run focused checks, the appropriate full suite, packaging checks and a fresh independent Sol/high review. The same worker handles accepted repairs. Copilot is user-disabled; no commit, PR or merge is authorized.

Local evidence and results are under `output/step3_gp_calculation_20261008/`. GPU implementation, all possible configurations and full-workbook Python recalculation remain outside this increment.
