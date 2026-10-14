# excel_to_act

English | [简体中文](README.zh-CN.md)

Use the complete [`excel-to-act` skill](.agents/skills/excel-to-act/SKILL.md) for agent-driven conversion. The [release guide](docs/release.md) covers the versioned skill/plugin ZIP, matching Python wheel, installation, checksums, and publication requirements. The runtime includes dedicated Step 6 delivery and Step 7 reporting skills.

Excel to Act converts Excel models into standalone Python tools through seven reviewed stages: extraction, indexing, input confirmation and design, code generation, validation, internal delivery, and a final conversion report. Start with `step1 report` after extraction is finalized. Each stage requires Agent review followed by an explicit human decision on its exact JSON/Markdown report pair.

## Workflow

```mermaid
flowchart TD
    X[Excel in input/] --> S1[1 · Extract and preserve]
    S1 --> A1{Agent → human}
    A1 -->|approve| S2[2 · Index and validate]
    S2 --> A2{Agent → human}
    A2 -->|approve| S3I[3a · Classify and confirm inputs]
    S3I --> A3I{Agent → human}
    A3I -->|confirm| S3[3b · Analyze and design]
    S3 --> A3{Agent → human}
    A3 -->|approve| S4[4 · Generate standalone Python]
    S4 --> A4{Agent → human}
    A4 -->|approve| S5[5 · Validate and reconcile]
    S5 --> A5{Agent → human}
    A5 -->|approve| S6[6 · Prepare internal package]
    S6 --> A6{Agent → actual human}
    A6 -->|approve| D[Deliver exact package and verify readback]
    D --> S7[7 · Final conversion report]
    S7 --> A7{Agent → actual human}
    A1 & A2 & A3I & A3 & A4 & A5 & A6 & A7 -->|reject| R[Return to a responsible stage · revise · review]
```

Use `step1 report --run FINAL_RUN --out NEW_DIR`, then follow the Step 2–5 commands and review gates. After Step 5 is approved, bind a matching comparison with `workflow bind-comparison --workflow NEW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256`; obtain the digest from a separate verification record. `workflow status --workflow DIR` shows current evidence, effective internal delivery, and the next permitted stage. Step 6 requires review of its exact package and report pair before delivery, followed by successful execution and readback of the delivered files. The [workflow guide](docs/workflow.md) also describes attaching approved parent evidence as read-only input.

## CLI and deliverables

Run `excel-to-act stepN tools`, `excel-to-act stepN agent`, or `--help` for current commands, limits, and instructions.

| Stage | Main commands | Main deliverables |
| --- | --- | --- |
| 1 · Extract | `step1 convert`, `check`, `finalize`, `report` | Preserved source parts, inventories, `import_checkpoint.json/.md` |
| 2 · Index and read | `step2 index`, `validate`, `prepare`, `query`, `trace`, `report` | Index, source-bound reading package, bounded evidence packets and traces |
| 3 · Analyze and design | `step3 prepare`, `fields`, `dependencies`, `input-catalog`, `query`, `trace`, `source-trace`, `produce-exact-cases`, `profile`, `plan`, `check`, `report`, `interface` | Confirmed inputs, targets, source candidates, variables, equations, modules, and `analysis_design.json/.md` |
| 4 · Generate | `step4 capture-external`, `discover`, `plan`, `generate` | Modular standalone bundle, source map, separate input/metadata records, generation report |
| 5 · Validate | `step5 validate`, `oracle`, `reconcile`, `finalize-cases` | Isolated execution evidence, native Excel capture, numerical comparison report |
| 6 · Internal package | `step6 candidate`, `deliver`, `revoke` | Versioned exact bundle, candidate JSON/Markdown pair, separate delivery and readback record |
| 7 · Final report | `step7 report`, `skill`, `template` | Human conversion report and JSON evidence |
| Workflow | `workflow attach`, `bind-comparison`, `status`, `confirm`, `reject` | Native or pinned-parent origin, current decisions, effective internal delivery |

Step 3 works backward from declared results and confirms scalar, vector, table, raw, and formula-derived input boundaries. Step 4 generates modular code only for supported, bound designs, after the semantic mapping and implementation preflight pass. Check the current tool catalog and source adapter for formula and input-shape limits; the workflow is not a universal Excel compiler. Formula caches are not model inputs, and extraction does not run VBA. Step 4 smoke evidence proves bundle execution only. Native Excel comparison requires Windows and Microsoft Excel. A tested scenario does not establish all-configuration coverage, actuarial certification, feedback solving, or GPU support.

Primary handoffs lead with the decision, scope, verified results, limits, unresolved choices, links to JSON/Markdown evidence, and the next review action. Detailed provenance and review records stay in the evidence appendix. Missing facts are unknown or “Not recorded,” never zero. `workflow.json` is the current acceptance ledger.

## Use

Python 3.11+ is required. Install the optional `vba` extra when using the VBA inspection tools.

The [Step 1 demo](examples/README.md) creates its own synthetic workbook and runs extraction through finalization. For a design with an `exact_cases` support domain, follow the [finite-case preparation](docs/cli.md#finite-case-preparation) and [32-case validation](docs/cli.md#exact-case-validation) routes. Each request needs its own validation and matching native capture; only `finalize-cases` over all 32 approved cases creates the formal Step 5 checkpoint.

```powershell
python -m pip install -e ".[vba]"
excel-to-act step1 tools
excel-to-act step1 agent
excel-to-act workflow status --workflow WORKFLOW_DIR
$BUNDLE = "output/WORKFLOW/stage4/bundles/revision-NNNN/bundle"
python "$BUNDLE/model.py" --inputs "REQUEST.json" --out "output/result.json"
```

## Development and packaging

Run these commands from the repository root. Clear the generated `build/` directory before every build so setuptools packages the current source and resources. Keep source files and workflow evidence outside `build/`.

```powershell
python -m pip install -e ".[dev]" build
python -m pytest -q
python -m ruff check .
if (Test-Path -LiteralPath .\build) { Remove-Item -LiteralPath .\build -Recurse -Force }
python -m build --outdir output/package-build
```

These development builds go to `output/package-build/`. For a versioned distribution, use the [release guide](docs/release.md), which requires a new output directory before building and preserves each completed package. Building a package does not approve or deliver a converted model.

See the [seven-stage workflow](docs/workflow.md), [CLI guide](docs/cli.md), [generic Agent contracts](docs/design/generic_agent_contracts.md), [result-driven analysis guide](docs/design/result_driven_analysis.md), and [report contract](docs/design/model_conversion_report.md). Source workbooks and run evidence under `input/` and `output/` are local data. The README pair is bilingual; other authored project documentation is in English.
