# examples

Small sanitized workbooks and expected artifact snapshots will live here. Do not commit client workbooks.

## Step1 directory demo

Use a directory of sanitized `.xlsx` / `.xlsm` fixtures as the human Step1 input. The CLI records unsupported and unreadable files alongside successful conversions:

```powershell
excel-to-act step1 tools
excel-to-act step1 agent
excel-to-act step1 convert .\examples\step1-raw --out .\step1-output
excel-to-act step1 check --run <run-dir>
excel-to-act step1 finalize --run <run-dir>
```

Use the `run_path` from `batch_handoff.json` (relative to the output directory) to select a per-source run. Inspect `handoff.md` for the human boundary and `handoff.json` for the host agent. Formula values are not recalculated. The agent chooses from the installed tool catalogue, follows diagnostics, and stops after three recovery attempts or the first no-progress result.

Logical coverage and package-part parsing use separate ratios. `--parsed-package-min` sets a lower bound on actually parsed package parts (default `0`), and `--opaque-max` sets an upper bound on opaque package parts (default `1`); each denominator is all non-directory OOXML ZIP members. Opaque parts are byte-preserved but never count as parsed. `--no-allow-opaque` rejects them regardless of those numeric limits.
