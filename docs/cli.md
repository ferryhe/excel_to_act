# CLI guide for the seven-stage conversion workflow

Start a seven-stage workflow with `step1 report --run FINAL_RUN_DIR --out WORKFLOW_DIR`. Each stage produces a JSON artifact and a Markdown handoff for Agent review followed by an explicit human decision. `workflow attach` requires a parent `workflow.json` with `schema_version: conversion.workflow.v1`, stage keys `1`–`6`, and `mode: production`. Its source, artifacts, and Stage 1–5 approvals must be current. Attach that evidence to a separate root:

```powershell
excel-to-act workflow attach --parent PARENT_DIR --out NEW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256
excel-to-act workflow status --workflow NEW_DIR
```

Supply `APPROVED_SHA256` from a separate comparison approval or verification record, independently of the file being attached. The attached root pins the exact parent manifest and comparison identity. Its inherited Stages 1–5 are read-only; its own Stages 6–7 require separate reviews.

Review the exact current JSON/Markdown pair as Agent first, then record the human decision only after an explicit user response. The Step 3 input catalog has its own review pair before the final design. TypeSafe delegation and live TypeSafe review commands are disabled. Do not create a receipt from an assumption, another report's approval, or a synthetic test.

Use a new workflow directory for a new conversion. The workflow manifest binds source files, stage reports, report revisions, and decisions. When an upstream report or input changes, later approvals become stale and must be reviewed again. `workflow status` shows the next permitted stage.

## 1. Review the source checkpoint

Complete and finalize Step 1 using the existing `step1` workflow. Then create a native seven-stage workflow with its first reviewable checkpoint over the final per-source run:

```powershell
excel-to-act step1 report --run FINAL_RUN_DIR --out WORKFLOW_DIR
excel-to-act workflow status --workflow WORKFLOW_DIR
```

Read `stage1/revision-N/import_checkpoint.json` and `import_checkpoint.md`. The report binds every file in the promotion ledger, `promotion.json`, the original workbook, and the preserved source copy. It shows measured quality, opaque parts, and extraction limits. It does not change or re-finalize the Step 1 run.

After review, record the Agent decision. Record the human decision only after the user explicitly approves or rejects the displayed checkpoint:

```powershell
excel-to-act workflow confirm --workflow WORKFLOW_DIR --stage 1 --reviewer agent --decision approve --message "Checked the source identity, complete promotion ledger, measured quality, and stated limits."
excel-to-act workflow confirm --workflow WORKFLOW_DIR --stage 1 --reviewer human --decision approve --message "<explicit user response>"
```

## 2. Review the saved index checkpoint

After Stage 1 is approved, validate and checkpoint the existing saved Step 2 index:

```powershell
excel-to-act step2 report --index INDEX_JSON --step1-root STEP1_ROOT --workflow WORKFLOW_DIR --source-id SOURCE_ID
```

The command runs the saved-index validator and records its source identity and reference checks. It does not rebuild the index or mutate Step 1/2 artifacts. Review `stage2/revision-N/index_checkpoint.json/.md`, then record separate Agent and human decisions with `--stage 2`. Stage 3 remains blocked until both decisions approve the current checkpoint.

## 3. Explore and report the design

After Stage 2 is approved, prepare a source-bound analysis and run the static workflow:

```powershell
$analysis = "output/conversion/analysis"
excel-to-act step3 tools
excel-to-act step3 prepare --index INDEX_JSON --step1-root STEP1_ROOT --out $analysis --source-id SOURCE_ID
excel-to-act step3 fields --analysis $analysis
excel-to-act step3 dependencies --analysis $analysis
excel-to-act step3 plan --analysis $analysis
excel-to-act step3 check --analysis $analysis
```

Use bounded queries and traces to record source facts and paths. These selectors are syntax examples only; replace them with the selected source's actual cells, fields, names or module identifiers:

```powershell
excel-to-act step3 query --analysis $analysis --target "Sheet1!A1" --out "output/conversion/packets/source-cell"
excel-to-act step3 query --analysis $analysis --target "Output!B2:B4" --offset 0 --limit 100
excel-to-act step3 query --analysis $analysis --target "RESULT_NAME"
excel-to-act step3 query --analysis $analysis --target "vba:MODULE" --out "output/conversion/packets/module"
excel-to-act step3 trace --analysis $analysis --target "RESULT_NAME" --direction upstream --max-depth 4 --max-fields 100
```

Before final semantic design, prepare an ordered target file and an Agent-authored `step3.input_boundary.input.v1` catalog. The catalog groups source scalars, vectors, and sparse tables, and separately accounts for axis metadata, source labels, constants, and unresolved retained coordinates. Use exact source ranges and coordinates; a sparse table must carry an explicit coordinate map. Do not infer that physical rows are business records or that adjacent columns share semantics. This phase describes the selected saved source only: `scenario.overrides` must be `{}`, and selectors are text metadata, not numeric values. Source, scenario, target, and catalog objects use closed documented field sets. Catalog `value_source_policy` keeps formula caches prohibited, names only the formula-mode source, and cannot claim formula-derived values are available now; future-source and binding declarations are non-empty text lists using the declared policy values. Object-form open questions require non-empty text `question_id` and `question`; any supplied source-record question reference must cite a declared question. Axis provenance descriptions and source header labels are text metadata. Unknown fields or wrong-typed payloads such as `sample_rates`, structured availability flags, `rates`, and numeric selector arrays are rejected before an artifact is appended. Source header labels remain verbatim in their documented metadata field. Formula-derived external values require a later, separately approved value-capture design; formula caches are never inputs.

```powershell
excel-to-act step3 input-catalog --analysis $analysis --workflow WORKFLOW_DIR --targets TARGETS.json --catalog INPUT_CATALOG.json
excel-to-act workflow status --workflow WORKFLOW_DIR
```

`input-catalog` may run without topology evidence. In that mode the result is a source-candidate catalog: target reachability, formula closure, dynamic branches, lookup-selected inputs, and input coverage remain unknown. A trace unbound to the current targets may be cited as a lead, but it does not prove the current boundary. Empty source cells omitted from stored fields must remain distinct from business inputs; blank-read behavior stays open if a formula trace shows a consumer. Review the exact `input_boundary.json/.md` pair and record both the Agent decision and the human's explicit response. A recommendation cannot supply either decision. Approval confirms the proposed catalog only and leaves Stage 3 next; it does not approve the final semantic design or release Step 4.

For a first-use workbook, preserve the accepted catalog's original `analysis.json` unchanged. Prepare a fresh continuation analysis using the same source, scope, fields, dependencies, and structural plan, then produce a source-only candidate trace and profile. This avoids changing the analysis file whose hash is bound by the catalog receipt:

```powershell
excel-to-act step3 prepare --index INDEX_JSON --step1-root STEP1_ROOT --out "output/conversion/analysis-continuation" --source-id SOURCE_ID --workflow WORKFLOW_DIR
$continuation = "output/conversion/analysis-continuation"
excel-to-act step3 fields --analysis $continuation
excel-to-act step3 dependencies --analysis $continuation
excel-to-act step3 plan --analysis $continuation
excel-to-act step3 source-trace --analysis $continuation --workflow WORKFLOW_DIR --out "output/conversion/candidate-trace"
excel-to-act step3 profile --analysis $continuation --source-trace "output/conversion/candidate-trace/source_candidate_trace.json" --out "output/conversion/source-profile"
```

`step3 source-trace` requires a currently confirmed input catalog and exact target list. It reads canonical Step 1 formula text and structure only. It unions statically readable conditional arms, stops at formula-derived external input extents, and reports raw catalog coordinates as reached, not reached in the known graph, or unresolved/possibly reached. Lookups do not identify a selected row or return value. Dynamic, broken, unsupported, excluded, and capped references remain explicit unknowns. The report is a source-candidate trace, not a Step 4 active trace; its candidate closure and active coverage remain unknown. It never reads formula caches, evaluates formulas, or calls Excel.

An Agent-authored finite coordinate fence can be supplied with `--model-scope INPUT.json`:

```powershell
excel-to-act step3 source-trace --analysis $continuation --workflow WORKFLOW_DIR --model-scope MODEL_SCOPE_JSON --out "output/conversion/scoped-candidate-trace"
```

The `gp.source_model_scope.input.v1` input copies the current candidate trace's `source`, `scenario`, `targets`, `analysis`, and `input_boundary` bindings. It also declares a hash-bound unscoped `full_candidate_trace`, finite `allowed_regions` with rationales, and separate `proposed_formula_starts` with rationales. The validator checks every target start and confirmed formula-derived external cut lies inside the fence; proposed starts must be current formula cells inside the fence and may not duplicate target starts. Whole-row/whole-column and larger finite reads are clipped to allowed regions before expansion. Each omitted rectangle is retained as a deduplicated frontier with its referring formula address/hash, original reference text, excluded region, and coordinate count. Candidate scope is labeled `scope_bounded_static_candidate_design`; omitted paths stay unknown and are not reported as unused.

`step3 profile --source-trace` writes `source_candidate_family_profile.json` and groups source formula candidates by syntax. A semantic map for this route must declare `selection_basis: static_candidate`, cite the candidate profile and trace, and represent every formula-derived external cut as role `external` with the exact confirmed `boundary_variable_id`, dimensions, and extents. These variables have no values and are not mapped as local formulas. For a scoped candidate trace, the semantic plan/check and Stage 3 report bind the scope file by path and SHA and preserve its scope label and frontier counts. Static plan/check coverage describes only the candidate set; it does not establish active closure or make Step 4 ready. A later Step 4 discovery must re-evaluate the approved target against current source formulas and prove the active subset before generation.

For a workbook that already has a current source-bound Step 4 trace, retain the existing active-selection profile route:

```powershell
excel-to-act step3 profile --analysis $analysis --trace ACTIVE_TRACE_JSON --out "output/conversion/active-profile"
```

Supply exactly one of `--trace`, `--source-trace`, or `--case-manifest` with `--case-id`. The active route profiles ordinary formulas in the supplied active trace; the candidate route profiles formulas reached from the confirmed input-catalog targets without requiring Step 4 history. Neither route reads trace values or formula caches, evaluates formulas, or infers business meanings. The finite-case route is described below.

`--target` accepts an exact name, finite cell/range, field ID, or canonical `vba:MODULE`. Use `--sheet` to resolve an unqualified address or local name. Query returns only stored cells; it does not fabricate blanks. A local name keeps its scope context for evidence replay. VBA queries return the canonical static inventory record; they do not execute VBA or infer worksheet ownership.

`query` pages default to 100 records and cap at 500. `trace` defaults to depth 4 and 100 fields and caps at 500 visited fields and 500 edge/boundary records. Frontier, unknown, excluded, and truncated boundaries remain explicit. An empty static frontier does not prove that `INDIRECT`, `OFFSET`, lookup-selected ranges, or conditional branches are closed; inspect the original formulas and record unresolved dynamic behavior.

`step3 profile --trace` writes `source_family_profile.json`; `step3 profile --source-trace` writes `source_candidate_family_profile.json`. Both compare formula text against the current canonical inventory and group translated syntax. Candidate arrays retain physical shape and follower evidence but still require explicit semantic shape mappings. Source-candidate counts must not be reported as active or numerically verified counts.

The optional `step3.model_spec.input.v1` draft has seven sections: `field_roles`, `logical_axes`, `calculation_groups`, `outputs`, `macro_workflow`, `open_questions`, and `validation_plan`. Use `fact`, `inferred`, and `pending` statuses. `details` can preserve structured targets, paths, axes, shapes, options, and implementation notes. `step3 check --spec INPUT.json` replays evidence selectors and validates exact source facts and references; it does not validate business meaning or numerical semantics.

Create an Agent-authored `step3.analysis_design.input.v1` file using the Step 3 agent contract. It must state targets and result ordering, source/scenario binding, conditions and paths, scope, known-only coverage, unresolved branches, implementation options, and a validation plan. Then write the Stage 3 report:

```powershell
excel-to-act step3 report --analysis $analysis --design DESIGN_INPUT_JSON --workflow WORKFLOW_DIR
```

Read the resulting `analysis_design.json/.md`, record Agent review, and wait for the human's explicit decision. For each `scenario.primary_inputs` entry that explicitly supplies both `source_cell` and `source_literal`, `step3 report` checks the qualified physical cell in the hash-bound Step 1 workbook copy with formulas visible. The cell must exist, contain no formula, and exactly match the declared literal. Partial assertions, missing cells, formula cells, and mismatches block the report before it appends a revision; diagnostics identify the scenario, selector, address, expected value, and actual value. The report records matched assertions under `scenario_source_validation`; it does not infer values or add undeclared guards. These static analysis and report commands do not calculate formulas, invoke native Excel, or read formula caches as runtime results. The separate `produce-exact-cases` route below calculates source traces from previously captured native evidence.

The design may include an optional `tool_execution_ledger` with text-only `tool`, `command`, `purpose`, `outputs`, and `status` fields. The report displays those records; it never executes listed commands. An optional `external_value_capture_plan` can describe the later phase, ranges, method, binding, and validation. It must say current values are unavailable and formula caches are prohibited. Step 3 records the prerequisite without capturing or including numeric values.

After the final design and input boundary have current approval, export the versioned interface:

```powershell
excel-to-act step3 interface --workflow WORKFLOW_DIR
```

Optional per-input declarations belong in the accepted design's `tool_interface.objects`, keyed by boundary `variable_id`. State `mutability` as `mutable`, `immutable`, or `unknown`; supply `value_type`, `required`, `default_policy`, `allowed_range`, `units`, and `unit_basis` only when supported by reviewed evidence. Default policy may be `unknown` or `none`; literal numeric defaults need later approved value capture. Missing facts are emitted as unknown and block runtime readiness. Boundary objects determine membership; explicitly mutable objects are included regardless of their business label. Parameter keys use the boundary `logical_name`; source addresses appear only in `source_provenance`. Export creates a manifest-listed derived artifact and Markdown report without adding a Step 3 revision or approval. CLI and workflow status say “interface generated” only while the registered artifact is current; they do not claim parameterized computation, Excel equivalence, or production readiness. The input boundary still requires Agent and actual human approval of its exact JSON/Markdown pair.

### Finite-case preparation

For an `exact_cases` design, prepare the complete request set after the input boundary is confirmed and before submitting the final design. The implemented validation route requires all 32 approved requests. Each case needs its own raw request, public request (`inputs` and same-request `external_data`), and previously captured native Excel oracle. Use the current source's approved files and SHA-256 values; neither another request's capture nor a saved-scenario result can stand in for them.

Supply a `step3.exact_cases.manifest.v1` JSON with a `cases` array. Each entry has exactly `case_id`, `raw_request`, `public_request`, and `native_oracle`; each file reference has `path` and `sha256`. The complete set must match the confirmed boundary's request hashes and case axes. Then seal and trace it:

```powershell
excel-to-act step3 produce-exact-cases --workflow WORKFLOW_DIR --cases CASE_MANIFEST.json --out NEW_CASES_DIR
```

The returned `artifact` is `NEW_CASES_DIR/exact_cases.json`. This command checks the existing native FullRebuild evidence and raw/public input readback, seals the formula-derived external cuts, and calculates a source formula trace for each request. It does not call Excel, use formula caches, or record a workflow decision. Use `step3 agent` for the source-bound case and semantic-map contracts.

For each case, create a distinct profile and Agent-authored semantic map, then plan and check it:

```powershell
excel-to-act step3 profile --analysis ANALYSIS_DIR --case-manifest NEW_CASES_DIR/exact_cases.json --case-id CASE_ID --out NEW_PROFILE_DIR
excel-to-act step3 plan --analysis ANALYSIS_DIR --semantic-map CASE_SEMANTIC_MAP.json --case-out NEW_PLAN_DIR
excel-to-act step3 check --analysis ANALYSIS_DIR --case-plan NEW_PLAN_DIR/semantic_plan.json
```

The semantic map must declare `selection_basis: exact_case` and bind this case's ID, manifest, trace, and profile by path and SHA-256. Add each case's checked profile and plan references alongside its request, producer, and trace references in the final design's `tool_interface.support_domain.cases`, then submit `step3 report` and obtain the normal final-design reviews. A profile or plan from the saved scenario cannot cover another request. After the current design is approved and its interface is exported, continue through Step 4 and the exact-case Step 5 route below.

## 4. Generate standalone Python

For each workflow, use its current confirmed input boundary and accepted semantic design. Approvals bind the exact design and bundle; changed evidence requires fresh reviews. Read source-specific parameters and manual checks from the current workflow's bound evidence, not from reusable defaults.

After the Stage 3 report has both current approvals, capture external inputs if required by that design, discover a fresh active trace, and pass implementation preflight before generation:

```powershell
excel-to-act step4 capture-external --workflow WORKFLOW_DIR
```

After the required capture, or directly when no capture is required:

```powershell
excel-to-act step4 tools
excel-to-act step4 discover --workflow WORKFLOW_DIR
excel-to-act step4 plan --workflow WORKFLOW_DIR --implementation IMPLEMENTATION_INPUT.json
excel-to-act step4 generate --workflow WORKFLOW_DIR
```

Step 4 generates modular bundles only. Follow the accepted capture and implementation design: run `step4 capture-external`, then `discover`, `plan --implementation INPUT.json`, and `generate` for the current catalog-bound modular path. Inspect `step4 tools` for supported boundaries and prerequisites. A design without its current semantic plan/check and implementation preflight is blocked; return to Step 3 to finish the modular design, then run the Step 4 plan. Normally omit `--trace` so generation uses the active trace already bound by the current implementation preflight. If supplied, it must be the exact latest passing Stage 4 discovery trace bound by that preflight; Stage 3 traces and older discovery revisions are rejected. Discovery evaluates the bound source formulas and raw values, never native Excel or cached formulas. Review `generation_report.json/.md` and its manifest-listed files: the entry point, logical equations/projection driver, modular runtime, raw/external/metadata ledgers and source map where required, together with the bounded isolated smoke evidence. Generation compiles and runs the bundle in a temporary directory without the workbook or project package; this proves execution, not equivalence with Excel. The bundle uses the Python standard library and reviewed source inputs, with no formula-cache outputs or project package import. Unsupported active formulas, cycles, external references, or unbound inputs block generation. Stage 5 becomes available after both confirmations under the active review policy.

## 5. Validate the generated code and reconcile with Excel

Run the generated code without the workbook or project package, then capture a new native Excel baseline and reconcile both results. Choose the saved-scenario or exact-case route from the current accepted support domain.

### Saved-scenario validation

The commands below validate the bundle's saved scenario. The target/range arguments are placeholders. Select all ranges and targets required by the current accepted design and manifest, and check the adapter's declared compatibility for the current source:

```powershell
excel-to-act step5 validate --workflow WORKFLOW_DIR
excel-to-act step5 oracle --workflow WORKFLOW_DIR --target RESULT_NAME --range "SHEET!A1:B10"
excel-to-act step5 reconcile --workflow WORKFLOW_DIR --validation VALIDATION_JSON --oracle EXCEL_ORACLE_JSON --abs-tol 1e-12 --rel-tol 1e-12
```

`step5 validate` compiles and runs the standalone generated bundle in an isolated temporary directory. It does not compare against Excel. `step5 oracle` requires Windows with Microsoft Excel, disables macros and events, recalculates a private read-only copy, captures requested defined names and finite ranges, then closes without saving. It records evidence only and does not advance Stage 5. `step5 reconcile` separately verifies the source, inputs, targets, active formula values, error codes, and tolerances. Both generated-code validation and reconciliation must pass before the Stage 5 report is reviewable. Review `validation_report.json/.md`, then record Agent and human decisions.

### Exact-case validation

For the current `exact_cases` contract, use the same sealed `exact_cases.json` that the approved Stage 3 design binds. Each public request must match its own trace, source, external data, and approved support domain. Run isolated validation with `--inputs`, then capture Excel with `--validation` pointing to that exact validation file, then reconcile those two outputs. The example uses `1e-12` absolute and relative tolerances; use the exact tolerances recorded in the approved support domain.

```powershell
$cases = (Get-Content -LiteralPath "EXACT_CASES.json" -Raw | ConvertFrom-Json).cases
$caseReports = @()
foreach ($case in $cases) {
    $validation = excel-to-act step5 validate --workflow WORKFLOW_DIR --inputs $case.public_request.path | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw "Validation failed: $($case.case_id)" }
    $oracle = excel-to-act step5 oracle --workflow WORKFLOW_DIR --validation $validation.validation | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw "Native capture failed: $($case.case_id)" }
    $comparison = excel-to-act step5 reconcile --workflow WORKFLOW_DIR --validation $validation.validation --oracle $oracle.artifact --abs-tol 1e-12 --rel-tol 1e-12 | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw "Reconciliation failed: $($case.case_id)" }
    $caseReports += @("--case-report", $comparison.artifact)
}
excel-to-act step5 finalize-cases --workflow WORKFLOW_DIR @caseReports
```

The returned `validation` from `validate` names the execution evidence file; `oracle.artifact` names the matching Excel capture; `reconcile.artifact` names that case's `case_report.json`. Each case receipt records only that request and does not advance Stage 5. `finalize-cases` requires 32 distinct current passing receipts and rechecks their source, request, trace, package, native result, and tolerance bindings. It writes one formal `validation_report.json/.md` pair for Agent review and actual-human acceptance. Saved-scenario evidence, fewer cases, or a successful report command cannot replace those decisions. Only after the formal checkpoint is approved should the workflow bind its comparison and proceed to Step 6.

Each reconciliation covers only its declared scenario, targets, formulas, and ranges. The workflow is not a general Excel compiler or an all-configuration model; a pass must not be described as broader than the evidence actually compared. Feedback solving or GPU execution requires separate implementation and evidence.

## 6. Package and internally deliver the approved tool

After native Stages 1–5 are current and approved, bind the independently verified comparison. Attached workflows already pin that comparison when they are created. `APPROVED_SHA256` must come from the separate verification record described at the start of this guide:

```powershell
excel-to-act workflow bind-comparison --workflow WORKFLOW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256
```

Then create a candidate for an exact request and matching native comparison:

```powershell
excel-to-act step6 tools
excel-to-act step6 candidate --workflow NEW_DIR --version VERSION --request REQUEST.json --comparison COMPARISON.json
excel-to-act workflow status --workflow NEW_DIR
```

Review the exact package JSON/Markdown pair and package ledger. Record an Agent decision, then obtain an explicit human decision on that pair with `workflow confirm --workflow NEW_DIR --stage 6 --reviewer agent|human --decision approve --message TEXT` (one reviewer per call). Smoke success or file presence is not acceptance. After current acceptance, run `excel-to-act step6 deliver --workflow NEW_DIR`. It copies the same package, executes its entry, checks the readback and records a separate effective internal delivery. `step6 revoke --workflow NEW_DIR --reason TEXT` makes it inactive while preserving package files and review records. This command provides internal delivery.

## 7. Create the final report

After current Step 6 acceptance and exact delivery/readback, run `excel-to-act step7 report --workflow NEW_DIR`. The report presents the accepted source, design, validation, provenance, and delivered package record. `step7 skill` and `step7 template` expose its packaged resources. Report generation creates a review pair for Agent and human decisions.

## Revisions and rejection

Use `excel-to-act workflow status --workflow NEW_DIR` to inspect stage freshness, decisions, effective delivery and the next allowed action. Reject a new package or report with its required return stage, for example:

```powershell
excel-to-act workflow reject --workflow NEW_DIR --stage 7 --reviewer agent --return-to 6 --message "The internal delivery evidence needs correction."
```

The rejected revision remains in history. A rejection routed backward invalidates the downstream path until a revised return-stage report is produced and approved. A changed source, input file, JSON report, Markdown report, or upstream revision also invalidates later approvals. Preserve old evidence; do not overwrite an accepted report or reuse its receipts for a replacement.
