# Progressive Excel Exploration for Agents

`step2 prepare`, `step2 query`, and `step2 trace` provide a progressive reading workflow for checked source evidence. This document records the reusable tool contracts, handoffs, limits, and validation expectations.

An agent should decompose a workbook once, verify the resulting facts, and then explore progressively through small evidence packets. Step 1 supplies source facts. Step 2 makes those facts accessible and preserves the approved scope. Step 3 interprets business behavior. Runtime observation and numerical reconciliation supply additional evidence when behavior depends on calculation or macros.

The Step 2 reading commands `prepare`, `query` and `trace` reuse extraction, indexing, views and graph code. `views validate` checks evidence packets. Human and machine handoffs are generated together from the same validated inputs. These tools do not add a formula engine, general agent orchestration framework, or complete VBA parser.

## Requirements

1. Reuse checked extraction artifacts instead of reparsing the raw workbook for every question.
2. Give an agent discoverable commands, bounded responses, source locations and explicit next actions.
3. Carry user scope decisions and read-only dependencies into both human and machine handoffs.
4. Preserve the distinction between source facts, static dependencies, interpretations and observed runtime results.
5. Support progressive exploration of workbook structures such as controls and VBA entry points when present.
6. Produce a concrete design that an independent reviewer can check against current code and these requirements.

## Current capabilities and limits

| Capability | Current behavior | Reusable reading rule |
| --- | --- | --- |
| Step 1 extraction | Directory conversion records raw facts, inventories, preserved package parts and fresh quality checks | Use the finalized source-bound handoff; keep optional graph and classification artifacts outside the factual quality gate |
| Control modules | Checkbox, ActiveX and VBA tool groups identify, convert and evaluate supported records; conversion supports dry-run | Treat links and event associations as static declarations; use their bound artifacts and do not claim execution |
| Step 2 index | Source and batch handoff resolution, reference validation, bounded recovery and resume | Preserve every source entry and its diagnostics; prepare a scope-bound reading bundle without changing source accounting |
| L1 views | Workbook, sheet and recorded-region views have stable record IDs, chunks and source locations | Read selected views and complete records; do not infer unrecorded cells or business meaning from layout |
| Formula graph | Static cell, name and supported table references with unresolved diagnostics | Trace supported references within explicit bounds and report frontiers or unresolved edges |
| VBA evidence | Readable module sources, declared procedures, event associations and optional literal cell links | Query module evidence as needed; dynamic targets and runtime behavior remain unresolved without separate evidence |
| Agent claims | Facts must exactly match supplied view records; derivations and unverified statements have separate labels | Validate claims against the delivered packet and its canonical source views |
| Paired handoffs | Human and machine packets bind to the checked source, selected scope and reading manifest | Keep both handoffs consistent with validated source, scope and reading evidence |

Primary implementation references:

- `src/excel_to_act/steps/step1/workflow.py`: extraction tools, fresh checks, source handoffs and promotion.
- `src/excel_to_act/steps/step1/agent.md`: source fidelity, bounded recovery and phase boundary.
- `src/excel_to_act/steps/step2/workflow.py`: reference resolution, validation, resume fingerprints and recovery history.
- `src/excel_to_act/steps/step2/agent.md`: indexing contract and downstream reading boundary.
- `src/excel_to_act/views/__init__.py`: deterministic views, record IDs, chunking and claim validation.
- `src/excel_to_act/graph/builder.py`: formula operand resolution and graph construction.
- `src/excel_to_act/inventory/vba_links.py`: literal VBA references and their known limits.
- `src/excel_to_act/interfaces/cli.py`: existing Typer command groups.
- `docs/design/agent_reading_contract.md`: existing source-bound claim contract.

The Step 2 index deliberately labels any non-empty usable index `partial`. `validation_status: pass` is the reference-integrity result. Preparation must not interpret the index usability label as a numerical or semantic assessment.

## Reference reading progression

| Stage | Agent action | Evidence produced | Condition for continuing |
| --- | --- | --- | --- |
| 1 Source intake | Identify source version and enumerate workbook structures | Source identity, manifests and feature capabilities | Supported source can be read; unsupported or failed inputs remain visible |
| 2 Source decomposition | Run existing Step 1 conversion and module extraction | Inventory, source facts, bindings, events, VBA sources and preserved package parts | Exact supported facts and package preservation pass existing checks |
| 3 Source acceptance | Run fresh check and finalize | Native source/batch handoff, quality and promoted output | Existing Step 1 `ready_for_next_step` is true |
| 4 Scope-aware preparation | Index accepted source; run Step 2 prepare | Reading manifest, views, dependency boundary and paired scoped handoffs | Source and reference integrity pass; scope is bound to the correct source |
| 5 Overview | Read source/tab summaries, feature capabilities and entry-point candidates | A small overview packet | Agent knows what can be queried and which scope boundaries apply |
| 6 Targeted exploration | Query cells, ranges, names, controls and VBA modules | Source-bound evidence packets with complete records and pagination | Enough evidence is delivered for the current question, or a specific limitation is recorded |
| 7 Dependency exploration | Trace the selected output, name or formula path | Bounded graph traversal, frontier and unresolved references | Agent distinguishes completed static traversal from truncation or dynamic uncertainty |
| 8 Semantic analysis | Describe the retained workflow and ask bounded TypeSafe questions where useful | Source-cited interpretations, module specifications and open questions | Step 3 conclusions remain reviewable and separated from facts |
| 9 Independent review | Review the specifications and evidence; request missing packets through the same tools | Review findings tied to task requirements | Accepted findings are resolved and reviewed before downstream implementation |

This is a reusable evidence progression, not a requirement to run every parser for every workbook. A workbook without VBA skips VBA interpretation. Prioritize opaque structures by their relevance to the requested workflow and confirmed scope; preserve unsupported or excluded structures as limits rather than silently adding them to downstream work.

## CLI surface

### Current command groups

The current CLI exposes `step1 tools`, `step1 agent`, `step1 convert`, `step1 check`, `step1 coverage`, `step1 fidelity`, `step1 finalize`, recovery tools, and separate identify/convert/evaluate groups for checkbox, ActiveX and VBA records. Step 2 exposes `index`, `validate`, `prepare`, `query`, `trace`, `report`, `tools`, and `agent`; `views compile` and `views validate` remain available for direct view workflows.

Step 1 continues to scan the full source. Scope decisions affect downstream work, not extraction denominators. Existing `step1.v1` and `step1.batch.v1` handoffs remain the native inputs to indexing. The companion `ai.task_handoff.v1` is an agent reading contract and is not passed to the current native index parser.

### Prepare command

```text
excel-to-act step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]
```

Preparation verifies the native index and referenced inputs, validates the source-bound scope, and compiles a reading bundle once. It writes a small manifest, separately addressable views, the graph needed for tracing, a dependency audit and cell snapshot derived from those checked source inputs, and the English handoffs. If no scope file is supplied, all source sheets are retained; preparation does not infer exclusions.

`--dry-run` returns the inputs, planned files and capability results without creating directories, output files or attempt-history entries. `--resume` reuses a bundle only when its source artifacts, scope, compiler version and options match and its saved outputs retain their expected hashes.

Preparation invokes a read-only integrity verifier underneath Step 2 rather than consuming another recovery dispatch. It must not fail merely because the existing three-attempt indexing recovery budget was already spent successfully.

### Dependency preparation contract

`--scope` is the user-decision input, not an AI handoff or a dependency snapshot. The source/run-bound `analysis.scope.v1` structure records source identity, retained/excluded sheets, decision authority and the saved-source dependency policy. Prepare checks identity against the selected native index entry, rejects unknown sheets or conflicting retained/excluded lists, and hashes the scope bytes. When a batch is supplied, this scope applies only to its matching source/run entry; other entries use the all-sheets default and remain visible in the manifest.

The dependency audit and snapshot are outputs, not external inputs. Prepare derives them from the native index's checksum-validated inventory, manifest/name declarations and supported static graph references. It uses the same inventory read already needed for views. For references from retained formulas into excluded sheets, it resolves supported direct A1 destinations and absolute named destinations, records every supported inbound reference and preserves the exact stored source cells and caches of the referenced ranges. It does not recalculate, fill absent cells or load the previous manually generated sidecars as authoritative values. Unsupported or dynamic references are reported as unresolved, so supported static counts do not assert complete dynamic coverage.

Legacy dependency filenames without bound hashes are hints, not trusted input files. Prepare does not infer authoritative dependencies from filenames or from a separate companion file. Imported or human-edited dependency values require an explicit source-bound, checksummed input contract.

Prepare copies the decision scope unchanged into the reading bundle. Generated audit/snapshot headers bind source SHA, run ID and that decision-scope hash. The reading manifest supplies the authoritative relative paths and SHA-256 values of both generated outputs; both human and AI handoffs refer to those manifest-bound outputs. Source-derived counts in the reports are computed from the checked inputs, not trusted from legacy scope summaries. There is no circular hash dependency: the snapshot names the decision-scope hash, while derived-file hashes live in the manifest rather than inside the decision scope.

If a required native source artifact is missing or changed, preparation blocks. If a referenced cache is unavailable, the snapshot reports it explicitly and does not claim numeric readiness. If a prepared audit or snapshot is later changed or missing, resume, read validation and handoff validation reject that bundle. Changing an unused legacy sidecar cannot change the generated snapshot. Source-derived counts must be recomputed from checked inputs and labeled with applicable static-resolution limitations.

### Query command

```text
excel-to-act step2 query --manifest MANIFEST --source-id SOURCE --kind KIND [--target TARGET] [--sheet SHEET] [--range A1_RANGE] [--budget N] [--cursor CURSOR] [--out PACKET]
```

Supported query kinds are `overview`, `sheet`, `cell`, `range`, `name`, `control`, `vba` and `feature`.

- `overview` lists source sheets, retained/excluded status, counts, available evidence and entry-point candidates from actual controls and VBA declarations.
- `sheet` returns a mechanical overview and a bounded first page, not every cell on the sheet.
- `cell` and `range` return stored values, formulas, cached values, availability and source locations. Missing stored cells remain absent; no synthetic zero is inserted.
- `name` resolves the declaration, scope and destination when supported. It does not choose an arbitrary match if the query lacks required sheet context.
- `control` selects canonical `checkbox_binding` or `activex_event` view records created during preparation from the checked module artifacts. It preserves invalid source links and records whether downstream use is excluded. Query-time presentation of a module artifact is not by itself a validated citation.
- `vba` returns the named module record and source-file reference. Initial delivery uses the complete module record; selecting a procedure does not silently rewrite the source fact into a snippet.
- `feature` explains which relevant structures are parsed as facts, preserved only, or available for a later specialist.

`--source-id` is required for query and trace, including a one-source bundle. The agent obtains it from the manifest's source list before calling either command. There is no implicit first-source selection. Ambiguous names return candidates and a `needs_selection` diagnostic. They never silently select the first workbook, sheet or definition.

### Scope checks on data selection

Query and trace enforce the reading scope bound by preparation, not just the agent's instructions. Overview and feature metadata can show every source sheet, its excluded status, counts and allowed dependency ranges. A sheet overview for an excluded sheet supplies this metadata without its ordinary cell/formula page.

For an excluded sheet, a direct cell query is eligible only inside a generated approved read-only dependency range. A range query must be fully contained in an approved range; an overlap alone is insufficient. A broader request returns a non-zero exit and structured `out_of_scope` diagnostic with the source, sheet, requested selector and allowed narrower targets. It does not silently clip the request or expose other data. If a name points into an excluded sheet, its destination data follows the same rule; an ineligible destination returns availability/scope metadata instead of full name-expression or cell facts.

Excluded controls and worksheet VBA modules supply availability metadata only, not binding interactions or code bodies for downstream analysis. Module ownership uses source-bound sheet/codeName evidence and the confirmed module mapping; unresolved ownership under an exclusion scope returns `needs_scope_resolution` rather than implicit eligibility. Retained controls and modules continue through the canonical evidence path.

Trace preserves the edge showing an excluded-sheet boundary and delivers only approved dependency data there. It stops expansion into that sheet's internal formulas, even when a cached dependency cell retains its formula text for provenance. A packet marks these records `read_only_dependency`; their source facts remain unchanged, but they are not a task to regenerate the excluded logic. A new confirmed scope and preparation revision are needed to broaden downstream reads. This is a reader-selection rule, not a change to source preservation or filesystem permissions.

The token budget is an estimate consistent with existing chunking. Records are indivisible. Oversized records are reported with their source-file reference and a suggested narrower query or larger budget; they are not truncated and then represented as exact source facts. A packet explicitly lists delivered record IDs, omitted or oversized records and the next cursor.

### Trace command

```text
excel-to-act step2 trace --manifest MANIFEST --source-id SOURCE --kind KIND --target TARGET [--sheet SHEET] --direction upstream|downstream|both [--max-depth N] [--max-nodes N] [--max-edges N] [--out PACKET]
```

The first trace targets are cells, ranges and defined names. Upstream means dependencies of the selected target; downstream means consumers. Existing formula edges run from the formula cell to its referenced input, so these meanings must be mapped deliberately rather than assuming conventional graph direction.

Trace resolves supported unambiguous absolute A1 name destinations and relates referenced ranges to their stored members within the limits. Formula members can continue the traversal. These resolution relations are derived structural relations with declaration/operand evidence; they are not source formula edges. Constant names can be terminal values. Relative, dynamic, external or unsupported multi-area definitions remain explicit unresolved boundaries.

Reverse traversal must also find a selected cell referenced through a containing range or a resolved name. Returning an empty consumer list after checking only exact cell-node edges would give an incomplete answer without saying so.

The traversal tracks visited identities, honors depth/node/edge limits, reports its frontier and sets `truncated` when limits prevent completion. Scope-boundary edges remain visible. Reading an approved range on an excluded sheet stops at the saved read-only data boundary; it does not reopen the sheet's internal conversion task.

Literal VBA references may be attached as evidence when available. They do not establish read/write direction, execution order or a complete call graph. Expressions such as `Range(Var1)` are unresolved dynamic targets. Interpretation of their possible values and workflow belongs to Step 3; the current literal extractor must not label them resolved automatically.

### Citation validation

```text
excel-to-act views validate --views EVIDENCE_PACKET --output CLAIMS
```

The validator supports the list-of-views format and the evidence-packet envelope. It verifies canonical view references and exact selected records before applying `validate_agent_output` rules. Only records delivered to the Agent are eligible citations; a record elsewhere in the workbook is not automatically permitted.

Facts retain the complete `facts` object and `SourceLocation`. A proposed interpretation remains a derivation or unverified claim. Passing this check establishes source citation consistency, not the correctness of the business interpretation or numerical result. Validation failure produces a structured diagnostic and a non-zero CLI exit.

### Canonical control records

Preparation reads the native index's checked `checkbox_bindings.json` and `activex_events.json` references. It validates their file hashes, typed models and workbook identity, then creates one `ViewRecord` per typed binding/control using the deterministic record helper. The record kinds are `checkbox_binding` and `activex_event`. Their `facts` are the complete serialized `CheckboxBinding` or `ActiveXControl`, including binding status and explicit source parts. Existing inventory control records remain available; reading records do not add objects to Step 1's logical accounting denominator.

Each canonical record belongs to its owner's sheet view. Its `SourceLocation` uses the recorded owner sheet, sheet part, shape identity and package-part fields, together with the validated source manifest. The binding/control object has no synthesized A1 address: its linked cell is a target in `facts`, not the control's source location. If an existing inventory source identity can be matched unambiguously, preserve it; otherwise use a stable declared control identity based on record kind, worksheet part and shape ID. Missing source metadata produces an explicit unavailable-provenance diagnostic rather than a guessed location.

The reading manifest binds these canonical views to the original control-artifact hashes. Query selects complete records from those views and lists their delivered IDs. Validation verifies both the module-artifact and canonical-view references and exact record contents before applying the existing claim rules. An altered binding, a binding from the wrong source/run, or a binding absent from the delivered packet fails. Static event association remains a source-derived declaration match and does not establish runtime event execution.

## Reading artifacts and reuse

A prepared bundle can use this small layout:

```text
reading/
  manifest.json
  scope.json
  dependency_audit.json
  retained_dependency_cells.json
  step1/HANDOFF.json
  step1/HANDOFF.md
  step2/HANDOFF.json
  step2/HANDOFF.md
  sources/<source-id>/views/<view-id>.json
  sources/<source-id>/dependency_graph.json
```

The source extraction artifacts and exported VBA files remain referenced instead of being copied again. Native source and promotion reports remain unchanged. A scope revision produces a new preparation revision; it does not edit checksum-bound final files or alter source measurements.

The manifest stores source/run identity, native index hash, scope hash, compiler version/options, per-view hashes and compact lookup entries. The typed scope and handoff/packet contracts validate their inputs and references. The current reading bundle uses a file-backed manifest and per-view JSON with the existing `WorkbookView`, `ViewRecord`, `SourceLocation` and graph models.

Compile and parse the source inventory once per preparation revision. Persist the already-built graph rather than building it again for each trace. A query loads only the required view files and verifies their hashes against the bound manifest; it does not reload unrelated views for every question.

Preparation performs full integrity validation. Each read verifies the current source identity, bound manifest and selected evidence hashes. Final handoff validation checks all referenced outputs. Hashing the small original workbook is different from reparsing every worksheet or regenerating all views; the former remains an identity check.

Bundle references use declared roots and relative paths so a receiving agent can resolve a relocated bundle. External roots, including the existing Step 1 artifact root, must be supplied explicitly. Absolute paths can be emitted as local convenience fields, but must not be the only portable identity. Original source-location fields remain exact provenance even when a local resolver uses a relocated file path.

## Machine contract

Both human and AI handoffs are rendered from one validated context, with these information groups:

- Native source and index references with SHA-256 and their path bases.
- Source quality and reference validation as separate states.
- Approved scope, decision authority, retained/excluded sheets and read-only dependencies.
- Available commands, argument contracts, capabilities, schema versions and read order.
- Current stage, completed work, remaining work and explicit next actions.
- Observed facts and their limits, without claiming macro execution or numerical equivalence.

An evidence packet adds selection metadata, canonical view references, the supplied `WorkbookView` slices, and pagination or traversal diagnostics. The slices preserve original view IDs, record IDs, record facts and source locations; selection does not create new source facts. Mechanical summaries and name/range expansion relations are labeled as derivations with their contributing source records.

All reader commands return structured status, identity, metrics, diagnostics and next actions. Absence of a supported match is distinct from ambiguity, missing artifacts, exclusion, truncation and unsupported resolution. Integrity errors block use. A bounded or partially resolved packet can remain usable, with `ready_for_static_analysis` scoped to the stated question and `workflow_verified: false` until behavior evidence exists.

Tool catalogues list implemented capabilities only. The Step 2 catalog's `reader_commands` section lists prepare/query/trace and references the view validator. Recovery-dispatch entries and their attempt history remain separate. Ordinary queries, cursor pages and traces do not consume the indexing recovery attempts.

## Reusable agent instructions

### Step 1 agent

Use the source-fidelity and recovery instructions in the packaged Step 1 Agent contract. When preparing a source handoff:

1. Read the live tool catalogue and use the declared module commands. Extract facts from the full source without inferring business scope.
2. Keep raw source fields authoritative. Separate missing caches, invalid links, preserved opaque parts and unsupported parsing from successful facts.
3. Check and finalize using existing gates. Never infer macro execution from source export or hide unsupported objects in the supported-fact denominator.
4. Provide a concise English human report and a machine-readable continuation containing source identity, artifact references, capabilities and the next stage.
5. Record user scope decisions as source-bound inputs. If the user changes scope after finalization, have preparation generate a new scoped reading handoff rather than editing final source artifacts.

### Step 2 reading agent

1. Read `step2 tools`, the native index and the prepared machine handoff. Confirm source and scope identity.
2. If no valid reading bundle exists, use prepare. Reuse a matching bundle instead of rerunning raw extraction.
3. Begin with overview. Choose a concrete question or entry-point candidate and query only the necessary names, cells, controls or VBA module.
4. Trace the selected dependency path with bounds. Read frontiers and unresolved diagnostics before claiming the path is complete.
5. At excluded-sheet boundaries, query only approved read-only ranges and retain their saved-scenario meaning.
6. Pass small source-cited packets to Step 3. Keep analytical conclusions in Step 3 output, not in the native index or source reports.
7. Validate citations against the delivered evidence. Missing evidence triggers another targeted read, not a guessed fact.
8. Stop repeating a query when its source, selector, scope and result are unchanged. Continue through a cursor or a different specific target only when it supplies new evidence. Escalate a concrete unresolved decision when local tools cannot answer it.

### Step 3 analysis and review

The main agent establishes inputs, outputs and workflow from evidence. TypeSafe/Jev latest supplies bounded typed judgments about interpretations, not exact lookups or arithmetic. Store the input evidence, question, typed response, uncertainty and resulting decision separately from source facts. A low-confidence or disputed interpretation prompts evidence retrieval or an explicit human question rather than silent acceptance.

A fresh independent reviewer receives the task requirements, current specifications and source packets. It checks realistic behavior and data-contract problems, including whether a claimed output path crosses a frozen dependency boundary or uses an unresolved VBA target. Findings require source evidence and a direct requirement mapping. The reviewer can request additional packets through the same reader tools. Source-citation validation is necessary but cannot replace semantic review.

## Source-backed exploration pattern

Start with an overview, then select a declared output, input, control or procedure from the current source. Query complete records with their original locations and trace supported static references within the confirmed scope. Treat control associations and VBA text as declarations, not proof of runtime execution. Keep dynamic destinations, unsupported references and excluded-sheet boundaries explicit. Changed-scenario claims require separate runtime evidence from a controlled workbook copy and a recorded comparison; they cannot be inferred from saved values or static source.

## Validation and operational checks

- Preparation rejects a wrong-source scope, missing native references and checksum mismatches without changing upstream artifacts. Derived outputs come from checked source data; unbound sidecars are not treated as authoritative.
- An unchanged reading bundle resumes. Changes to source, scope, compiler inputs or generated artifacts invalidate the affected reuse.
- Query packets contain complete source records, original locations, explicit pagination and only the selected eligible records. Ambiguous names require enough context; missing cells remain distinct from zero.
- Workbook-local and worksheet-local names, batch source selection, oversized records and control/module records follow their declared CLI contracts.
- Tracing resolves only supported name and range references, finds range consumers, preserves edge direction, and terminates at configured limits. Dynamic, unsupported and excluded-scope boundaries remain explicit.
- Citation validation rejects changed or wrong-run facts, tampered references, and records that were not delivered in the packet.
- Paired reports share provenance and labels. Static analysis and source citations do not establish business correctness, runtime behavior or numerical equivalence.
- Reader operations reuse prepared views and graph data. Record performance counts and timings only when they were measured for the current source and run.
