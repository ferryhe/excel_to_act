# Step 3: Evidence-based workflow analysis

Step 3 interprets the current workbook's requested workflow using checked source evidence. It follows Step 3 input confirmation and produces a reviewable design; it does not calculate formulas, run macros, generate code or prove numerical equivalence.

## Entry and scope

Start only after `workflow status` shows the current Step 2 evidence and decisions are valid. Bind analysis to the selected source/run, current scope decision, target order and confirmed input boundary. Preserve source facts and upstream artifacts. Recompute source-derived evidence from checked current artifacts; an earlier report or dependency snapshot is a lead, not proof of the current boundary. If scope or an input assumption is materially unclear, ask the user and keep the design pending.

## Evidence and interpretation

Use the source-bound views and `step3 query` to inspect relevant cells, ranges, names, controls and modules. Use `step3 trace` for bounded static paths and record unresolved or dynamic references. Distinguish:

- source facts and original source locations;
- structural derivations such as range membership or dependency edges;
- business interpretations and user-confirmed choices;
- runtime observations and numerical comparisons, when separately captured.

Describe input/output mappings, operation order, loop and recurrence rules, state changes, conditions, error handling, and shared calculations only when supported by evidence. Do not treat a current table entry or cached formula value as proof of every possible target or changed-scenario result. Retain unknowns and explain which evidence or decision is still needed.

## Handoff contents

The paired `analysis_design.json/.md` should bind the source, scope, scenario, targets and input-boundary revision. It should include:

- requested outputs and declared result order;
- confirmed inputs grouped by logical kind with source locations and any recorded axes or units;
- supported calculation paths, modules, dependencies, conditions and recurrence order;
- unresolved references, scope limits, unsupported behavior and open business questions;
- implementation alternatives and their scoped coverage;
- execution and comparison evidence still required before later stages.

Use known-only denominators when evidence is incomplete and label the scope explicitly. Keep planned, generated, smoke-tested and native-compared work separate. The Step 3 `check` command validates its declared structural and evidence contracts; it does not establish business correctness or numerical readiness.

## Review checks

- Every source fact matches current, supplied evidence and the correct source/run.
- Cell coordinates support provenance; they are not automatically logical business objects.
- Target order, shape, input kind, axes, units, and recurrence are recorded only when supported.
- Dependencies crossing a scope boundary, dynamic references, opaque structures, and capped traces remain visible.
- Material ambiguities remain questions until the user supplies a decision.
- The report and JSON preserve facts, derivations, interpretations and unknowns separately.
- Step 3 planning does not imply generation, execution, equivalence or acceptance.
