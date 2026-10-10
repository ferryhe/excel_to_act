# Current architecture and evidence boundaries

The public conversion path is a six-stage workflow. Each stage reads source-bound inputs, emits machine-readable evidence and a paired human handoff, and is reviewed through the workflow ledger. The stage-local `agent.md` defines the Agent contract; `tools/` provides the executable catalog used by `excel-to-act stepN tools`.

Implementation paths in the tables are relative to `src/excel_to_act/` unless they begin with another repository path.

## Stage responsibilities

| Stage | Main implementation | Inputs and outputs | Boundary |
| --- | --- | --- | --- |
| 1 · Extract and preserve | `steps/step1/workflow.py`, `source_scan.py`, `checkpoint.py` | Workbook files become source facts, inventories, preserved package parts, quality checks and a source handoff/checkpoint consumed by Step 2. | Records source content and measured extraction limits. It does not assign business meaning, recalculate formulas, or run macros. |
| 2 · Index and read | `steps/step2/workflow.py`, `prepare.py`, `query.py`, `trace.py`, `checkpoint.py`; `views/` and `graph/` | A checked Step 1 handoff becomes a source-bound index and reading bundle. Query and trace return bounded evidence packets consumed by Step 3. | Performs deterministic navigation and supported static reference tracing, not semantic interpretation or runtime verification. |
| 3 · Analyze and design | `steps/step3/workflow.py`, `input_boundary.py`, `exploration.py`, `calculation.py`, `design.py`, `semantic.py`, `profiling.py` | Uses the requested results and accepted upstream evidence to describe targets, input boundaries, source paths, variables, equations, modules, and execution order. The accepted design and catalog feed Step 4. | Produces design and static analysis evidence; it does not calculate formulas or establish numerical equivalence. Input-boundary confirmation remains a distinct human decision. |
| 4 · Generate | `steps/step4/discovery.py`, `implementation.py`, `generator.py`, `modular_bundle.py` | Consumes the accepted design and catalog, checks supported implementation coverage, and emits a standalone Python bundle with source maps and generation evidence for Step 5. | A bounded bundle smoke check proves that the generated package runs in isolation. It is not a native Excel comparison; supported formula and adapter limits remain source-specific. |
| 5 · Validate and reconcile | `steps/step5/workflow.py` and `steps/step5/native_excel_oracle.ps1` | Validates the generated bundle, captures a fresh native Excel result when available, and compares declared outputs and supported evidence against the current design. Its reports feed Step 6. | Standalone execution, native calculation, and comparison are separate evidence. Only recorded comparison scope supports a reconciliation claim; the commands do not record approval. |
| 6 · Report | `steps/step6/workflow.py` and the packaged Step 6 report skill/template | Consumes accepted upstream evidence and emits the reader-facing final report and machine-readable report pair for Agent and human review. | Does not calculate formulas or change earlier evidence. The workflow ledger remains the authority for formal acceptance. |

## Cross-stage services

- `interfaces/cli.py` connects user commands to the stage workflows. Each `steps/stepN/tools/` package supplies its current tool catalog, and each stage's `agent.md` describes how to use it.
- `steps/conversion_workflow.py` stores stage revisions, upstream freshness, review decisions, and return routes in the source-bound workflow. A successful command or report does not itself approve a stage.
- `schemas/` defines shared artifact contracts. `store/` and stage-local writers persist evidence and provenance. Human report renderers present evidence; they do not create new formula or comparison results.
- Data moves from preserved source facts to indexed evidence, reviewed design, generated code, validation results, and the final handoff. A downstream interpretation does not alter its source evidence.

The standalone `excel-to-act inspect` path remains available for exploratory analysis and its own optional numerical checks. It is separate from the six-stage approval ledger and cannot substitute for a current Step 1–6 workflow.

## Evidence distinctions

Source formulas and saved caches are distinct facts. Static references do not prove runtime behavior. Step 3 design is planning evidence; Step 4 smoke evidence is generated-code execution; Step 5 native capture and reconciliation are separate, source- and target-bound evidence. A passed technical check is not formal acceptance, and a tested scenario does not certify every configuration or workbook feature.
