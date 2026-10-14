# excel_to_act

English | [简体中文](README.zh-CN.md)

Start a new seven-stage workflow with `step1 report`, or attach one to a current approved six-stage legacy workflow. A native workflow reviews its own Stages 1–5; an attached workflow inherits those stages read-only from the exact pinned legacy parent. Legacy Stage 6 remains its historical final report. New Step 6 packages the approved standalone tool for internal delivery, and Step 7 produces the final conversion report. The CLI reports legacy status but does not write to legacy workflows. The Step 3 input boundary still requires an actual human decision.

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
    A5 -->|approve| S6[6 · Package internal candidate]
    S6 --> A6{Agent → actual human}
    A6 -->|approve| D[Deliver exact package and verify readback]
    D --> S7[7 · Final conversion report]
    S7 --> A7{Agent → actual human}
    A1 & A2 & A3I & A3 & A4 & A5 & A6 & A7 -->|reject| R[Return to a responsible stage · revise · review]
```

For a new conversion, use `step1 report --run FINAL_RUN --out NEW_DIR`, then the existing Step 2–5 commands and Agent/human review gates. After its own current approved Step 5, use `workflow bind-comparison --workflow NEW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256`. Supply that digest from a separate verification record. For migration, use `workflow attach --parent LEGACY_DIR --out NEW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256` only after legacy Stages 1–5 and their source are current. The attached root pins the exact legacy manifest and comparison identity. `workflow status --workflow DIR` shows current evidence, effective internal delivery, and the next permitted stage. A candidate smoke or report generation does not approve a stage. The candidate JSON/Markdown pair needs Agent review and an explicit human decision before delivery; old Stage 6 receipts or delegation cannot approve it. Rejected and delivered history remains in the ledger.

## CLI and deliverables

Run `excel-to-act stepN tools`, `excel-to-act stepN agent`, or `--help` for current commands, limits, and instructions.

| Stage | Main commands | Main deliverables |
| --- | --- | --- |
| 1 · Extract | `step1 convert`, `check`, `finalize`, `report` | Preserved source parts, inventories, `import_checkpoint.json/.md` |
| 2 · Index and read | `step2 index`, `validate`, `prepare`, `query`, `trace`, `report` | Index, source-bound reading package, bounded evidence packets and traces |
| 3 · Analyze and design | `step3 prepare`, `fields`, `dependencies`, `input-catalog`, `query`, `trace`, `source-trace`, `profile`, `plan`, `check`, `report` | Confirmed inputs, targets, source candidates, variables, equations, modules, and `analysis_design.json/.md` |
| 4 · Generate | `step4 capture-external`, `discover`, `plan`, `generate` | Modular standalone bundle, source map, separate input/metadata records, generation report |
| 5 · Validate | `step5 validate`, `oracle`, `reconcile` | Isolated execution evidence, native Excel capture, numerical comparison report |
| 6 · Internal package | `step6 candidate`, `deliver`, `revoke` | Versioned exact bundle, candidate JSON/Markdown pair, separate delivery and readback record |
| 7 · Final report | `step7 report`, `skill`, `template` | Human conversion report and JSON evidence |
| Workflow | `workflow attach`, `bind-comparison`, `status`, `confirm`, `reject` | Native or pinned-parent origin, current decisions, effective internal delivery |

Step 3 works backward from declared results and confirms scalar, vector, table, raw, and formula-derived input boundaries. Step 4 generates modular code only for supported, bound designs, after the semantic mapping and implementation preflight pass. Check the current tool catalog and source adapter for formula and input-shape limits; the workflow is not a universal Excel compiler. Formula caches are not model inputs, and extraction does not run VBA. Step 4 smoke evidence proves bundle execution only. Native Excel comparison requires Windows and Microsoft Excel. A tested scenario does not establish all-configuration coverage, actuarial certification, feedback solving, or GPU support.

Primary handoffs lead with the decision, scope, verified results, limits, unresolved choices, links to JSON/Markdown evidence, and the next review action. Detailed provenance and history stay in the evidence appendix. Missing facts are unknown or “Not recorded,” never zero. `workflow.json` is the current acceptance ledger.

## Use

Python 3.11+ is required. Install the optional `vba` extra when using the VBA inspection tools.

```powershell
python -m pip install -e ".[vba]"
excel-to-act step1 tools
excel-to-act step1 agent
excel-to-act workflow status --workflow WORKFLOW_DIR
$BUNDLE = "output/WORKFLOW/stage4/bundles/revision-NNNN/bundle"
python "$BUNDLE/model.py" --inputs "REQUEST.json" --out "output/result.json"
```

See the [seven-stage workflow](docs/workflow.md), [CLI guide](docs/cli.md), [generic Agent contracts](docs/design/generic_agent_contracts.md), [result-driven analysis guide](docs/design/result_driven_analysis.md), and [report contract](docs/design/model_conversion_report.md). Source workbooks and run evidence under `input/` and `output/` are local data. The README pair is bilingual; other authored project documentation is in English.
