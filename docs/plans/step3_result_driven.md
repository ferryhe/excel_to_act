# Step 3: Result-driven analysis and design

Step 3 starts from declared outputs and scenario scope, then traces the required calculations backward through current source-bound evidence. Use the current workflow's targets, inputs, equation families, axes, recurrences, and implementation boundaries. This document gives reusable analysis guidance; it contains no workbook-specific values or acceptance conclusions.

## Analysis sequence

1. Declare each output selector, its order, shape, meaning when known, and scenario.
2. Confirm the input boundary separately. Classify logical scalars, vectors, tables, raw values, formula-derived external values, metadata and unresolved inputs. Record source locations without treating cell coordinates as business objects.
3. Use `step3 query` and `step3 trace` to collect complete records and follow supported static references. Read frontiers, exclusions, dynamic references and missing evidence before describing a path as complete.
4. Describe conditional paths, shared calculations, errors, equations, module responsibilities, axes, initialization and recurrence order from the source and accepted decisions.
5. Record alternatives, external-input boundaries, known-only coverage, unresolved interpretations, and validation work. Keep planned, generated, smoke-tested and Excel-compared results separate.
6. Prepare the paired `analysis_design.json/.md`, check its source bindings and evidence, and hand it to the Agent and human review gates.

The Step 3 CLI supports source preparation, field and dependency summaries, source queries and traces, source-candidate profiles, structural planning and checks. The [CLI guide](../step3_cli.md) describes current commands and data contracts. These tools navigate and validate evidence; they do not calculate formulas or establish Excel equivalence.

## Target and path model

A target may be a named result, a cell, a finite range, or a field ID. For example, `Output!B2:B4` can describe an ordered vector target when that meaning is recorded. Declare shape and axis metadata from the request or source-bound evidence; do not infer a business time axis from cell geometry. Whole-row and whole-column selectors are not bounded queries; first identify a finite extent.

For each supported path, record the target mapping, conditions, required source inputs, shared modules, calculation order, error behavior, and evidence. Count shared calculations once. Dynamic names, unresolved lookups, conditional behavior, unsupported formulas, opaque components, and frontier limits remain unknown until evidence resolves them. A static trace does not prove runtime execution.

## Implementation options and coverage

Describe scope choices without deciding for the user. Options may include a complete set of supported configurations, a fixed declared scenario, staged modules, or an explicit external-input boundary. State the unit, numerator, denominator, scenario and stage for each coverage measure. Keep documented, source-resolved, planned, generated, smoke-tested and native-compared coverage distinct. A complete known set does not prove that all semantic paths or supported configurations are known.

For formula-derived external inputs, identify the producer, boundary variable, shape/axis contract, value source, capture method and downstream responsibility. Formula caches and prior results are not substitute inputs. If the source adapters or supported formula subset cannot execute the declared design, record that gap rather than claiming readiness.

## Review checks

- Targets, source, scenario, input catalog and upstream artifacts are bound to the current workflow.
- Each factual claim is traceable to supplied evidence; derivations and business interpretations are labeled separately.
- Query and trace limits, missing records, dynamic behavior, excluded scope and unresolved boundaries remain visible.
- Input kinds, axes, units, target order and recurrence are stated only when recorded or confirmed.
- Alternatives and unresolved business choices are explicit; a clarification is not an approval.
- Planning is not presented as calculation, generation, runtime evidence or numerical reconciliation.
- JSON and Markdown agree, and the current workflow ledger—not the report's technical status—records acceptance.
