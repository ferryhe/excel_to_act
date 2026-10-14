# Excel to Act skill release

The distribution pairs the complete seven-stage `excel-to-act` skill with the Python runtime from the same source. It supports local agent execution and reviewed internal model conversion. Source workbooks, workflow receipts, native Excel captures, and delivered business models stay in their conversion directories.

The repository skill lives at `.agents/skills/excel-to-act/`. The Python package also includes dedicated Step 6 delivery and Step 7 reporting skills. `stepN agent`, `stepN tools`, and `workflow status` define the executable workflow and its human review gates.

## Build the distribution

Use Python 3.11+ and run from the repository root. Start with an empty release output directory. Keep local evidence under `input/` and `output/` and out of packaging intermediates.

```powershell
$releaseDir = "output/release-0.1.0"
if (Test-Path -LiteralPath $releaseDir) { throw "Release directory already exists; choose a new path." }
New-Item -ItemType Directory -Path $releaseDir -ErrorAction Stop | Out-Null
python -m pip install -e ".[dev]" build
python -m pytest -q
python -m ruff check .
if (Test-Path -LiteralPath .\build) { Remove-Item -LiteralPath .\build -Recurse -Force }
python -m build --outdir $releaseDir
python tools/build_skill.py --wheel "$releaseDir/excel_to_act-0.1.0-py3-none-any.whl" --out $releaseDir
```

The directory guard runs before either build command. Each build uses a fresh output directory and preserves completed distributions. `build_skill.py` checks the distribution name/version and every packaged Python module and declared resource against the current source. A stale, missing, or changed file blocks packaging, and existing release ZIPs and checksum files are not overwritten. The package version and wheel path must match.

The resulting files are:

| File | Purpose |
| --- | --- |
| `excel_to_act-0.1.0.tar.gz` | Python source distribution |
| `excel_to_act-0.1.0-py3-none-any.whl` | Installable Python CLI and stage resources |
| `excel-to-act-skill-0.1.0.zip` | Portable plugin layout with one complete skill and its matching wheel |
| `SHA256SUMS` | SHA-256 of those release artifacts |

Inside the ZIP, `excel-to-act/plugin.json` identifies the portable plugin, `skills/excel-to-act/` is the installable skill folder, and `release.json` records the version, `packaged` status, license presence, runtime path, and exact packaged file hashes. The wheel is under that skill's `runtime/`. Packaging selects only the maintained skill files and verified wheel. The MIT `LICENSE` is included in both the plugin and standalone skill folders, and the Python wheel and source distribution include the same license text.

## Install and verify locally

Extract the ZIP to a new directory. Verify each release artifact against `SHA256SUMS` with `Get-FileHash -Algorithm SHA256` before using it. Keep installed skill folders and their matching runtime wheel together.

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

Publish from a reviewed commit with passing CI. Retain its exact distributions, checksum list, local installation result, and review evidence. Check that the ZIP and wheel identify the same version and that all seven packaged Agent contracts are readable. CI retains the distribution as `excel-to-act-distribution`. Creating a GitHub Release, uploading to PyPI, or submitting a plugin marketplace listing is a separate publishing action.

The project is licensed under the [MIT License](../LICENSE). Python package metadata declares `License-Expression: MIT`, and the skill manifest records `license_recorded: true`. Dependency licenses retain their own terms. Optional `oracle-formulas` is a separate EUPL opt-in and is not bundled into the default runtime.

The validated real-model scope is the Pricing workbook and its 32 approved requests. Each conversion's accepted evidence defines its supported source and scenarios. Public distribution, production activation, and publication of source workbooks or conversion evidence require explicit instructions for those actions.
