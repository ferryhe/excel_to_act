# examples

The demo below creates a small synthetic workbook locally. Do not commit client workbooks.

## Step1 directory demo

Run from the repository root after installing the package. Create the demo under the ignored `input/` directory, then convert, check, and finalize it under `output/`. The CLI also records unsupported and unreadable files alongside successful conversions:

```powershell
excel-to-act step1 tools
excel-to-act step1 agent
$demoInput = "input/step1-demo"
$demoOutput = "output/step1-demo"
New-Item -ItemType Directory -Path $demoInput -Force | Out-Null
python -c "from openpyxl import Workbook; import sys; book = Workbook(); sheet = book.active; sheet.title = 'Demo'; sheet['A1'] = 2; sheet['A2'] = 3; sheet['B1'] = '=SUM(A1:A2)'; book.save(sys.argv[1]); book.close()" "$demoInput/example.xlsx"
$converted = excel-to-act step1 convert $demoInput --out $demoOutput | ConvertFrom-Json
$run = Join-Path $demoOutput $converted.entries[0].run_path
excel-to-act step1 check --run $run
$finalized = excel-to-act step1 finalize --run $run | ConvertFrom-Json
$finalRun = Join-Path $demoOutput $finalized.final_output
$finalRun
```

The conversion JSON gives each candidate's `entries[].run_path`, relative to the output directory. The saved `batch_handoff.json` is inside the returned batch artifact path. Finalization returns `final_output`; use `$finalRun` as `step1 report --run` for a new workflow. Inspect `handoff.md` for the human boundary and `handoff.json` for the host agent. Formula values are not recalculated, and finalization does not record Agent or human approval. The agent chooses from the installed tool catalogue, follows diagnostics, and stops after two consecutive no-progress results or three cumulative recovery attempts.

Logical coverage and package-part parsing use separate ratios. `--parsed-package-min` sets a lower bound on actually parsed package parts (default `0`), and `--opaque-max` sets an upper bound on opaque package parts (default `1`); each denominator is all non-directory OOXML ZIP members. Opaque parts are byte-preserved but never count as parsed. `--no-allow-opaque` rejects them regardless of those numeric limits.
