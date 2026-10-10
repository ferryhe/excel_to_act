# GP scope: WaiverOfPrem disabled

The user has deferred WaiverOfPrem for the current GP conversion. The unchanged saved workbook already satisfies this decision: the six premium-waiver payout selectors at `Main!F40:F45` are the text `None`. Their fixed amounts at `Main!D40:D45` are zero. No source edit or replacement Python bundle is needed.

| Item | Source location | Current treatment |
| --- | --- | --- |
| Six waiver payout selectors | `Main!F40:F45` | `None` for every category |
| Selected benefit column | `Benefit_Table!F3`, `F5:F110` | `None`; 106 literal zero payouts |
| GP-dependent alternative | `Benefit_Table!K3`, `K5:K110` | `WaiverOfPrem`; excluded from the active GP calculation |
| Waiver payout matrix | `Premium!BK10:BP115` | 106 years × 6 categories; all zero in Python and the recorded fresh native Excel capture |
| Waiver cost vector | `Premium!BX10:BX115` | 106 zero costs in Python and the recorded fresh native Excel capture |
| Conditional feedback path | `P_GP_FEEDBACK_CONDITIONAL` | Inactive; no feedback solver |

The alternative formula starts at `Benefit_Table!K5`:

```text
=-PV(PricingIntRate,MAX(PremPayPeriod-C5,0),GP)
```

Selecting it can introduce `GP -> WaiverOfPrem -> PVFB -> GP`. The separate source cost equation, `BX[t] = SUMPRODUCT(BK:BP, AN:AS) * T[t]`, stays in the model and its reconciliation. Retaining this zero-cost path preserves source/category alignment without activating GP feedback. ROP also references GP and remains unselected.

The existing Step 4 saved-scenario adapter binds these selectors and rejects changes to them. Enabling WaiverOfPrem later requires a revised Step 3 scope and equations, an explicit feedback-solving design, regenerated code and new reconciliation. The source workbook's iterative settings are retained as oracle evidence; this decision does not change Excel's global calculation settings.

The source-bound [verification record](../../output/gp_source_analysis_20261009/waiver_disabled_scope-r1.json) checks the source selectors, zero payouts/costs, active dependencies and existing adapter guard. It reuses the current Step 4 revision-6 bundle and Step 5 attempt-3 model/native capture; it does not launch Excel or use formula caches as model inputs. The [human-readable check](../../output/gp_source_analysis_20261009/waiver_disabled_scope-r1.md) lists the evidence and result.

This scope clarification preserves the existing model and numerical evidence. Current GP remains `2.7283284514524744`. Steps 1–6 are accepted for the saved case. Step 5 revision 6 has Agent and actual-human acceptance. The [human-facing Step 6 report](../../output/stage36_20261008/stage6/revision-0004/conversion_report.md) has independent PASS, Agent approval and live delegated TypeSafe acceptance with approve probability 0.87, meeting the unchanged 0.85 gate. The [workflow](../../output/stage36_20261008/workflow.json) remains the approval ledger.
