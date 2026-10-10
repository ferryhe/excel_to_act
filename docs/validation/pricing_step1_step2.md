# Pricing: verified Step 1 → Step 2

Result: PASS on 2026-10-08 using the current working tree and default quality thresholds. The local workbook was extracted, checked, finalized, indexed and validated. Its bytes remain unchanged.

## Result

| Item | Final evidence |
| --- | --- |
| Source | `input/Pricing.xlsm`, 2,618,416 bytes |
| SHA-256 | `dc1de98b5e27fa85e233e7ccfa56bc21f571b2bb1686ab13b86a43a9cdfe7c7c` |
| Source run | `20261008T120940-8282f510` |
| Workbook | 49 sheets, 157,764 cells, 101,451 formula cells; every formula cell has a saved cache |
| Names | 549 declarations in the manifest; 550 expanded destination records in inventory |
| Supported objects / facts | 160,952 accounted; 160,952 exact; both ratios 100% |
| Package bytes | 625/625 parts preserved; 86 parsed, 539 opaque |
| Controls / VBA | 454 checkboxes: 189 resolved, 265 invalid; 8 ActiveX controls with 8 declared handlers; 50 VBA modules |
| Step 1 gate | `partial`, `ready_for_next_step: true`, no blockers; finalized under the default opaque policy |
| Promotion | All 691 recorded file checksums match; source-evidence copy matches the original |
| Step 2 | Single-source and batch indexes each retain one entry and 62 artifact references; `validation_status: pass` |
| Resume | `index.json`, `INDEX.md` and `state.json` unchanged byte-for-byte; no new attempt |
| Tests | 171 passed in 53.43 seconds; 19 dependency deprecation warnings |

`partial` records the source limitations; it is not a failed integrity check. Step 2 also labels a non-empty usable index `partial`. Initial indexing checked all 62 artifact references; unchanged validation and resume reused that checked evidence.

## Local deliverables

Output root: `output/real_steps_20261008/`.

- [Human run report](../../output/real_steps_20261008/run_report.md)
- [Machine report and verified paths](../../output/real_steps_20261008/run_report.json)
- [Single-source index](../../output/real_steps_20261008/step2/INDEX.md)
- [Batch index](../../output/real_steps_20261008/step2_batch/INDEX.md)
- [Resume hashes](../../output/real_steps_20261008/resume_verification.json)
- [Step 2 commands and results](../../output/real_steps_20261008/step2_execution.json)

`logs/` holds CLI JSON responses and stderr. Candidate and accepted output directories are listed in the machine report. Workbooks and runtime outputs remain local and are ignored by Git.

## Reproduce

See the current [CLI contracts](../../README.md#cli-and-deliverables); `step1 tools`, `step2 tools` and `--help` describe executable arguments. Conversion writes the candidate and initial checks; `finalize` performs another fresh check before promotion. This run also invoked `step1 check` explicitly, exercised both source and batch handoffs, and resumed the source index.

After indexing, repeat its command with `--resume`. Use the same handoff, Step 1 root and output directory. For batch indexing, select the `batch_handoff.json` reference returned by conversion after finalizing eligible entries.

The focused test command was:

```powershell
python -m pytest -q tests/test_step1.py tests/test_step2_cli.py tests/test_step2_index.py tests/test_control_artifacts.py tests/test_views.py
```

## Step 3 boundary

This run proves source extraction, supported-fact fidelity, preserved bytes, promotion and index integrity. It does not execute macros or independently recalculate formulas. It did not compile a real-workbook views bundle or implement Step 3.

The [Step 3 plan](../plans/step3_analysis_plan.md) starts with the checked `PremiumTable.CmdPremium_Click` source path. It retains the saved four-sheet exclusion decision, whose source hash matches this run; its older dependency audit must be regenerated for a new analysis bundle. Fresh retained-sheet checks found 49 resolved checkbox bindings and no invalid links.
