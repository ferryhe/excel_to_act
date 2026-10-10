# Pricing conversion case and manual acceptance

This page contains source-specific evidence for `input/Pricing.xlsm`. It is an example of applying the generic six-step workflow, not a set of defaults for another workbook. Source SHA-256 is `dc1de98b5e27fa85e233e7ccfa56bc21f571b2bb1686ab13b86a43a9cdfe7c7c`.

## Source-specific design facts

- `Premium!H1:H3` are three scalar summaries; `Premium!H10:H115` is one projection series with literal seed `H10` and recurrence `H11:H115`.
- H/I/J cover years 1–106. T covers years 0–106, including the time-zero seed at `Premium!T9`. Their lengths and array starts differ.
- The source's prior-H reference in the I recurrence, I-squared CI-cost term, and first-period versus later-period payment error handling are preserved and explained in the model guide.
- The confirmed catalog has 44 logical inputs: 18 scalars, 24 vectors and two sparse tables. Five formula-derived CI curves are explicitly accepted as external inputs: `Qtable!W5:W110`, `Y5:Y110`, `Z5:Z110`, `AD5:AD110` and `AH5:AH110`, keyed by `Qtable!C5:C110`. Their upstream algorithms are outside this implementation.
- Seven numerical paths are implemented and reconciled. The GP-linked feedback condition is inactive. The user deferred WaiverOfPrem; `Main!F40:F45` are `None`, and waiver payouts and annual costs are zero. See the [waiver scope](../plans/gp_waiver_scope.md) and its source-bound checks.
- The present Stage 5 primary-input adapter checks seven `Main` cells: C3, C4, C6, C7, C8, J48 and O26. This is a current implementation restriction; another workbook requires its own reviewed input/capture/reconciliation bindings.

Read the [confirmed inputs](../../output/stage36_20261008/stage3/revision-0022/input_boundary.md), [accepted design](../../output/stage36_20261008/stage3/revision-0025/analysis_design.md), [model guide](../../output/gp_source_analysis_20261009/step4_reading_guide-r4.md) and [validation report](../../output/stage36_20261008/stage5/revision-0006/validation_report.md) for exact evidence and scope. Original source facts and accepted report pairs remain unchanged.

## Current completion state

Steps 1–6 are accepted for the unchanged saved GP case. Step 5 revision 6 has Agent and actual-human acceptance. The [human-facing Step 6 report](../../output/stage36_20261008/stage6/revision-0004/conversion_report.md) has independent Sol/high PASS, Agent approval and live delegated TypeSafe acceptance with approve probability 0.87, meeting the unchanged 0.85 gate. It explains inputs, calculation order, measured results, limits and run instructions. The [workflow ledger](../../output/stage36_20261008/workflow.json) is authoritative; this page grants no approval.

## Manual acceptance checks

First review the business boundary: the result is GP for the saved source scenario; CI curves are inputs; WaiverOfPrem is disabled. Confirm that scalar/series/table classification, time alignment and module explanations match the intended actuarial model. The existing input and design approvals need no duplicate confirmation merely to reread them. A different desired boundary or source correction requires a revised design and affected reviews.

For an optional independent run, extract the [portable bundle](../../output/gp_source_analysis_20261009/delivery/gp_python_saved_case.zip) into an empty directory. With Python 3.11+, run these commands from that directory:

```powershell
python model.py --out gp-review.json
(Get-Content -Raw gp-review.json | ConvertFrom-Json).targets
```

No project installation or Excel is required for that run. Compare these four roots:

| Target | Expected saved-case value |
| --- | ---: |
| GP | 2.7283284514524744 |
| AnnuityDue | 8.574676381865864 |
| PVLoading | 3.161998467798486 |
| PVFB | 14.767563151498457 |

The technical report already compares all 8,507 active formula identities and values, 530 external values/source formula cuts and 106 age keys against a fresh native Excel calculation, with zero mismatches at absolute/relative tolerance 1e-12. Manual Excel checking is optional: use a copy, retain the saved parameters and compare `Premium!J1` and `H1:H3`. Do not save over the original source used by the approval ledger.

Do not treat changes to age, term, switches or raw tables as tests the saved-case bundle promises to support. Its guard deliberately rejects changed inputs. Such scenarios require a newly reviewed design/bundle and reconciliation. GPU execution and feedback solving have not been certified.

Formal stage acceptance is separate from testing. An explicit actual-human approval restores Agent+human only for that current, unrejected artifact pair; other stages retain their existing delegation. Step 6 revision 4 has passed the delegated gate. The optional manual checks above do not require duplicate stage confirmation; a changed source, input boundary or implementation requires the affected revisions and reviews to be renewed.
