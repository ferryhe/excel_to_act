# Excel to Act skill release

The release candidate pairs the complete seven-stage `excel-to-act` skill with the Python runtime from the same source. It is intended for local agent execution and reviewed internal model conversion. Source workbooks, workflow receipts, native Excel captures, and delivered business models are not part of this software distribution.

The repository skill lives at `.agents/skills/excel-to-act/`. The earlier Step 6 and Step 7 skills remain stage-specific resources inside the Python package; they are not the complete conversion entry point. The total workflow and its human gates remain defined by `stepN agent`, `stepN tools`, and `workflow status`.

## Build a candidate

Use Python 3.11+ and run from the repository root. Start with an empty release output directory. Keep local evidence under `input/` and `output/` and out of packaging intermediates.

```powershell
$releaseDir = "output/release-candidate-0.1.0"
if (Test-Path -LiteralPath $releaseDir) { throw "Candidate directory already exists; choose a new path." }
New-Item -ItemType Directory -Path $releaseDir -ErrorAction Stop | Out-Null
python -m pip install -e ".[dev]" build
python -m pytest -q
python -m ruff check .
if (Test-Path -LiteralPath .\build) { Remove-Item -LiteralPath .\build -Recurse -Force }
python -m build --outdir $releaseDir
python tools/build_skill.py --wheel "$releaseDir/excel_to_act-0.1.0-py3-none-any.whl" --out $releaseDir
```

The directory guard runs before either build command. Keep the old directory intact and choose another path for every replacement candidate, even when its version is unchanged. `build_skill.py` also checks distribution name/version and every packaged Python module and declared resource against the current source. A stale, missing, or changed file blocks packaging, and existing release ZIPs and checksum files are not overwritten. Future versions must update the package version and matching wheel path together.

The resulting files are:

| File | Purpose |
| --- | --- |
| `excel_to_act-0.1.0.tar.gz` | Python source distribution |
| `excel_to_act-0.1.0-py3-none-any.whl` | Installable Python CLI and stage resources |
| `excel-to-act-skill-0.1.0.zip` | Portable plugin layout with one complete skill and its matching wheel |
| `SHA256SUMS` | SHA-256 of those release artifacts |

Inside the ZIP, `excel-to-act/plugin.json` identifies the portable plugin, `skills/excel-to-act/` is the installable skill folder, and `release.json` records the version, candidate status, license presence, runtime path, and exact packaged file hashes. The wheel is under that skill's `runtime/`. Packaging selects only the maintained skill files and verified wheel; it does not collect the working directory. A root `LICENSE`, when present, is retained in both the plugin and standalone skill folders.

## Install and verify locally

Extract the ZIP to a new directory. Verify each release artifact against `SHA256SUMS` with `Get-FileHash -Algorithm SHA256` before using it. Do not replace an existing skill installation without choosing how to retain that version.

```powershell
Expand-Archive -LiteralPath "$releaseDir/excel-to-act-skill-0.1.0.zip" -DestinationPath output/skill-install-check
$skill = (Resolve-Path "output/skill-install-check/excel-to-act/skills/excel-to-act").Path
python -m venv output/skill-runtime-check
& output/skill-runtime-check/Scripts/python.exe -m pip install "$skill/runtime/excel_to_act-0.1.0-py3-none-any.whl"
& output/skill-runtime-check/Scripts/excel-to-act.exe --help
& output/skill-runtime-check/Scripts/excel-to-act.exe step7 template
```

Python dependencies may need network access during installation; the bundle is not a complete offline dependency mirror. VBA inspection additionally needs the `vba` extra. The runtime installs no skill automatically and requires no OpenAI API key. Native Excel comparison needs Windows with Microsoft Excel; generated model bundles run without Excel.

For repository use, Codex discovers `.agents/skills/excel-to-act/`. For use in another project, copy the extracted complete skill folder, including `runtime/`, into that project's `.agents/skills/` or the user's skill directory, then invoke `$excel-to-act`. The extracted root also follows the [official portable plugin format](https://developers.openai.com/plugins/build/plugins); installation in a plugin host and marketplace review are separate from building this ZIP. See the [official skill discovery guidance](https://learn.chatgpt.com/docs/build-skills) for supported local locations.

Test the installed runtime with the [synthetic Step 1 demo](../examples/README.md), using that environment's Python and CLI. Verify that it stops for actual-human stage decisions. Installation, package smoke tests, and software publication must not create conversion receipts or deliver a business model.

## Release decision

Finalize from the reviewed PR commit after its current CI passes. Rebuild and retain the exact distributions, checksum list, local install result, and review evidence for that commit. Check that the ZIP and wheel identify the same version and that all seven packaged Agent contracts are readable. CI builds and retains the candidate artifacts for review; it does not tag, create a GitHub Release, upload to PyPI, or submit a plugin marketplace listing.

Before public distribution, the owner must choose and record the project's license. No project license is inferred from the default dependency license check. A candidate without `LICENSE` has `license_recorded: false`; it is not declared ready for unrestricted public reuse. Optional `oracle-formulas` remains a separate EUPL opt-in and is not bundled into the default runtime.

The accepted real-model evidence covers the original Pricing workbook and the original 32 requests described in [PR 49](https://github.com/ferryhe/excel_to_act/pull/49). This software candidate does not broaden that support domain or certify other workbooks. Public distribution, production activation, and publication of source workbooks or conversion evidence require their own explicit instruction; this release preparation does not perform them.
