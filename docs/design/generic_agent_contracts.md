# Generic Agent contracts and source-specific evidence

The six packaged `agent.md` files define reusable actions, evidence requirements, deliverables and review gates. Each is paired with a stage-local `tools/` catalog package that backs `stepN tools`; `stepN agent` presents the stage instructions. For each conversion, the Agent uses the current source and user request to select tools, explore checked evidence, inspect limits, and prepare paired JSON and Markdown handoffs. Source facts and case-specific choices belong in that workflow's evidence under `output/`, never in reusable instructions.

Before entering a stage, inspect `workflow status` and verify that the stage is permitted and its upstream evidence and decisions are current. A draft command, successful report creation, or clarification does not itself approve or advance a stage. Legacy standalone source-exploration APIs remain available for discovery; they do not create workflow approvals.

The default review order is Agent first and actual human second on the exact current JSON/Markdown pair. A new human or TypeSafe approval requires matching current Agent approval. Use TypeSafe only under an already registered, explicitly scoped authorization; an approval from another source or run cannot be reused. The Step 3 input-boundary checkpoint always requires actual human review. Ask the user when targets, scope, input kind, axes, units, or business interpretation remain materially uncertain; clarification is not a final stage decision. Rejections preserve history, return to the responsible current or earlier stage, and require revised affected reports plus fresh reviews.

| Information | Where it belongs |
| --- | --- |
| Workflow actions, provenance, error handling, stage deliverables and review gates | Packaged stage Agent contracts |
| Targets, source locations, selector values, input groups, shapes, axes and exclusions | Current source-bound catalog and design artifacts |
| Source exceptions, selected conditions and expected numerical results | The current workflow's evidence and reports |
| Accepted scenario guards, source-to-variable mappings and executable equations | The generated bundle and its manifest |
| Actual formula subset, supported capture shapes and adapter restrictions | Tool capability contracts and implementation; check compatibility for each source |
| Illustrative source facts or manual acceptance instructions | A clearly identified run-specific output, not reusable Agent defaults |

Schema identifiers and manifest-listed filenames are tool interfaces, not business defaults. Their names do not authorize a target, worksheet, scenario, or reused decision. Reusable instructions describe the process; the current catalog and source-bound evidence determine what a particular conversion can support.

## Verification and limits

Use `stepN agent` to read the installed contract and `stepN tools` to inspect capabilities and limits. Check links and current package resources when changing shared instructions. This contract does not establish that every Excel workbook can complete Steps 4–5: supported formulas, external capture, and source-specific adapters require compatibility checks and may require additional implementation. A completed workflow for one source does not certify another source or scenario.
