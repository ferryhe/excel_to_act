# Progressive Excel Exploration for Agents

Status: proposed design. The new commands described below are not implemented. This analysis covers tools, CLI contracts, agent instructions and acceptance checks; independent review is required before treating the design as ready for implementation.

An agent should decompose a workbook once, verify the resulting facts, and then explore progressively through small evidence packets. Step 1 supplies source facts. Step 2 makes those facts accessible and preserves the approved scope. Step 3 interprets business behavior. Runtime observation and numerical reconciliation supply additional evidence when behavior depends on calculation or macros.

The minimum addition is three Step 2 commands: `prepare`, `query` and `trace`. They reuse the existing extraction, indexing, views and graph code. The existing `views validate` command gains support for evidence packets. English human and machine handoffs are generated together from the same validated inputs. No new formula engine, general agent orchestration framework or complete VBA parser is needed for this scope.

## Requirements

1. Reuse checked extraction artifacts instead of reparsing the raw workbook for every question.
2. Give an agent discoverable commands, bounded responses, source locations and explicit next actions.
3. Carry user scope decisions and read-only dependencies into both human and machine handoffs.
4. Preserve the distinction between source facts, static dependencies, interpretations and observed runtime results.
5. Support progressive exploration of the real Pricing workbook, including its retained checkbox bindings and VBA entry points.
6. Produce a concrete design that an independent subagent can review against current code and these requirements.

## Current capabilities and gaps

The current runtime already has useful foundations. The missing layer is the repeatable reading workflow around them.

| Capability | Current behavior | Required addition |
| --- | --- | --- |
| Step 1 extraction | Directory conversion, raw source facts, inventory, package preservation and independent checks | Retain the existing gates; add a concise agent continuation section to generated handoffs |
| Control modules | Separate checkbox, ActiveX and VBA identify, convert and evaluate CLI groups; conversion supports dry-run | Expose these as discoverable module capabilities and reference their existing artifacts in reading packets |
| Step 2 index | Source and batch handoff resolution, reference validation, bounded recovery and resume | Prepare a scope-aware reading bundle without changing source accounting |
| L1 views | Workbook, sheet and recorded-region views; stable record IDs, chunks and source locations | Write separately addressable view files and return selected records without loading every view |
| Formula graph | Static cell references, named references and supported table references; unresolved diagnostics | Bounded traversal plus literal name/range expansion and explicit incomplete results |
| VBA evidence | Readable module sources, declared procedures, event associations and optional literal cell links | Query individual modules and expose unresolved dynamic targets; semantic workflow interpretation remains Step 3 |
| Agent claims | Facts must exactly match supplied view records; derivations and unverified statements have separate labels | Validate claims against the actual delivered packet and its canonical source views |
| Scoped AI handoffs | Companion `HANDOFF.json` files were assembled for the current run | Generate and schema-check them from validated source, scope and reading manifests |

Current source locations for implementation assessment:

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

## Fixed progression

| Stage | Agent action | Evidence produced | Condition for continuing |
| --- | --- | --- | --- |
| 1 Source intake | Identify source version and enumerate workbook structures | Source identity, manifests and feature capabilities | Supported source can be read; unsupported or failed inputs remain visible |
| 2 Source decomposition | Run existing Step 1 conversion and module extraction | Inventory, source facts, bindings, events, VBA sources and preserved package parts | Exact supported facts and package preservation pass existing checks |
| 3 Source acceptance | Run fresh check and finalize | Native source/batch handoff, quality and promoted output | Existing Step 1 `ready_for_next_step` is true |
| 4 Scope-aware preparation | Index accepted source; run proposed Step 2 prepare | Reading manifest, views, dependency boundary and paired scoped handoffs | Source and reference integrity pass; scope is bound to the correct source |
| 5 Overview | Read source/tab summaries, feature capabilities and entry-point candidates | A small overview packet | Agent knows what can be queried and which scope boundaries apply |
| 6 Targeted exploration | Query cells, ranges, names, controls and VBA modules | Source-bound evidence packets with complete records and pagination | Enough evidence is delivered for the current question, or a specific limitation is recorded |
| 7 Dependency exploration | Trace the selected output, name or formula path | Bounded graph traversal, frontier and unresolved references | Agent distinguishes completed static traversal from truncation or dynamic uncertainty |
| 8 Semantic analysis | Describe the retained workflow and ask bounded TypeSafe questions where useful | Source-cited interpretations, module specifications and open questions | Step 3 conclusions remain reviewable and separated from facts |
| 9 Independent review | Review the specifications and evidence; request missing packets through the same tools | Review findings tied to task requirements | Accepted findings are resolved and reviewed before downstream implementation |

These are fixed stages, not a requirement to run every parser for every workbook. A workbook without VBA skips VBA interpretation. Opaque structures are prioritized by their relevance to a retained path. Excluded checkbox repairs do not enter the current conversion queue.

## CLI surface

### Commands retained

Keep `step1 tools`, `step1 agent`, `step1 convert`, `step1 check`, `step1 coverage`, `step1 fidelity`, `step1 finalize`, the recovery tools, all nine control-module commands, `step2 index`, `step2 validate`, `step2 tools`, `step2 agent`, `views compile` and `views validate`.

Step 1 continues to scan the full source. Scope decisions affect downstream work, not extraction denominators. Existing `step1.v1` and `step1.batch.v1` handoffs remain the native inputs to indexing. The companion `ai.task_handoff.v1` is an agent reading contract and is not passed to the current native index parser.

### Proposed prepare command

```text
excel-to-act step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]
```

Preparation verifies the native index and referenced inputs, validates the source-bound scope, and compiles a reading bundle once. It writes a small manifest, separately addressable views, the graph needed for tracing, a dependency audit and cell snapshot derived from those checked source inputs, and the English handoffs. If no scope file is supplied, all source sheets are retained; preparation does not infer exclusions.

`--dry-run` returns the inputs, planned files and capability results without creating directories, output files or attempt-history entries. `--resume` reuses a bundle only when its source artifacts, scope, compiler version and options match and its saved outputs retain their expected hashes.

Preparation invokes a read-only integrity verifier underneath Step 2 rather than consuming another recovery dispatch. It must not fail merely because the existing three-attempt indexing recovery budget was already spent successfully.

### Dependency preparation contract

`--scope` is the user-decision input, not an AI handoff or a dependency snapshot. The initial supported input is the source/run-bound `analysis.scope.v1` structure already used by Pricing: source SHA/run identity, retained/excluded sheets, recorded decision authority and the saved-source dependency policy. Prepare checks identity against the selected native index entry, rejects unknown sheets or conflicting retained/excluded lists, and hashes the scope bytes. When a batch is supplied, this scope applies only to its matching source/run entry; other entries use the all-sheets default and remain visible in the manifest.

The dependency audit and snapshot are outputs, not external inputs. Prepare derives them from the native index's checksum-validated inventory, manifest/name declarations and supported static graph references. It uses the same inventory read already needed for views. For references from retained formulas into excluded sheets, it resolves supported direct A1 destinations and absolute named destinations, records every supported inbound reference and preserves the exact stored source cells and caches of the referenced ranges. It does not recalculate, fill absent cells or load the previous manually generated sidecars as authoritative values. Unsupported or dynamic references are reported as unresolved, so supported static counts do not assert complete dynamic coverage.

The current Pricing scope contains legacy `dependency_policy.snapshot` and `audit` basenames. Prepare treats them as historical hints and reports that they were not consumed. It does not infer trusted input files from those basenames or follow the hashes in a separate AI companion. Imported or human-edited dependency values are outside this initial contract. A future import requirement would need an explicit source-bound, checksummed input contract.

Prepare copies the decision scope unchanged into the reading bundle. Generated audit/snapshot headers bind source SHA, run ID and that decision-scope hash. The reading manifest supplies the authoritative relative paths and SHA-256 values of both generated outputs; both human and AI handoffs refer to those manifest-bound outputs. Source-derived counts in the reports are computed from the checked inputs, not trusted from legacy scope summaries. There is no circular hash dependency: the snapshot names the decision-scope hash, while derived-file hashes live in the manifest rather than inside the decision scope.

If a required native source artifact is missing or changed, preparation blocks. If a referenced cache is unavailable, the snapshot reports it explicitly and does not claim numeric readiness. If a prepared audit or snapshot is later changed or missing, resume, read validation and handoff validation reject that bundle. Changing an unused legacy sidecar cannot change the generated snapshot. Pricing acceptance requires deriving the existing 864 supported inbound references, 12 ranges and 1,464 stored records from the checked source, with the static-resolution limitations retained.

### Proposed query command

```text
excel-to-act step2 query --manifest MANIFEST --source-id SOURCE --kind KIND [--target TARGET] [--sheet SHEET] [--range A1_RANGE] [--budget N] [--cursor CURSOR] [--out PACKET]
```

Supported initial kinds are `overview`, `sheet`, `cell`, `range`, `name`, `control`, `vba` and `feature`.

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

### Proposed trace command

```text
excel-to-act step2 trace --manifest MANIFEST --source-id SOURCE --kind KIND --target TARGET [--sheet SHEET] --direction upstream|downstream|both [--max-depth N] [--max-nodes N] [--max-edges N] [--out PACKET]
```

The first trace targets are cells, ranges and defined names. Upstream means dependencies of the selected target; downstream means consumers. Existing formula edges run from the formula cell to its referenced input, so these meanings must be mapped deliberately rather than assuming conventional graph direction.

The current graph ends at many name or range nodes. A usable trace must resolve unambiguous absolute A1 name destinations and relate referenced ranges to their stored members within the limits. Formula members can then continue the traversal. These added resolution relations are derived structural relations with declaration/operand evidence; they are not new source formula edges. Constant names can be terminal values. Relative, dynamic, external or unsupported multi-area definitions remain explicit unresolved boundaries.

Reverse traversal must also find a selected cell referenced through a containing range or a resolved name. Returning an empty consumer list after checking only exact cell-node edges would give an incomplete answer without saying so.

The traversal tracks visited identities, honors depth/node/edge limits, reports its frontier and sets `truncated` when limits prevent completion. Scope-boundary edges remain visible. Reading an approved range on an excluded sheet stops at the saved read-only data boundary; it does not reopen the sheet's internal conversion task.

Literal VBA references may be attached as evidence when available. They do not establish read/write direction, execution order or a complete call graph. Expressions such as `Range(Var1)` are unresolved dynamic targets. Interpretation of their possible values and workflow belongs to Step 3; the current literal extractor must not label them resolved automatically.

### Extend existing validation

```text
excel-to-act views validate --views EVIDENCE_PACKET --output CLAIMS
```

Keep support for the existing list-of-views format. Add normalization of the new packet envelope. Packet validation verifies its canonical view references and exact selected records before using the existing `validate_agent_output` rules. Only records actually delivered to the agent are eligible citations; a record elsewhere in the workbook is not automatically permitted.

Facts retain the complete `facts` object and `SourceLocation`. A proposed interpretation remains a derivation or unverified claim. Passing this check establishes source citation consistency, not the correctness of the business interpretation or numerical result. Validation failure produces a structured diagnostic and a non-zero CLI exit.

### Canonical control records

Preparation must extend the view compiler to read the native index's checked `checkbox_bindings.json` and `activex_events.json` references. It validates their file hashes, typed models and workbook identity, then creates one `ViewRecord` per typed binding/control using the existing deterministic record helper. The record kinds are `checkbox_binding` and `activex_event`. Their `facts` are the complete serialized `CheckboxBinding` or `ActiveXControl`, including the binding status and explicit source parts. Existing inventory control records remain available; new reading records do not add objects to Step 1's logical accounting denominator.

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

The manifest stores source/run identity, native index hash, scope hash, compiler version/options, per-view hashes and compact lookup entries. Reuse existing `WorkbookView`, `ViewRecord`, `SourceLocation` and graph models. Add a typed scope and handoff/packet envelope only where the current ad hoc structures need validation. A file-backed manifest and per-view JSON are sufficient initially; a database is justified only by measured query costs.

Compile and parse the large source inventory once per preparation revision. Persist the already-built graph rather than building it again for each trace. A query loads only the required view files and verifies their hashes against the bound manifest. It does not deserialize the entire 773 MB Pricing views file for every question.

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

Tool catalogues list implemented capabilities only. Step 2 adds a `reader_commands` section for prepare/query/trace and references the view validator. Existing recovery-dispatch entries and their attempt history remain separate. Ordinary queries, cursor pages and traces do not consume the three indexing recovery attempts.

## Proposed agent instructions

### Step 1 agent

Retain the existing source-fidelity and recovery instructions. Add these continuation requirements when the new reading workflow is available:

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

## Pricing acceptance walkthrough

The existing run provides a realistic pilot: 45 retained sheets, 49 valid retained checkbox bindings, 8 matched ActiveX events and 50 exported VBA modules. Four excluded sheets contain all 265 invalid checkbox links. Retained formulas still have 864 static references into 12 preserved ranges on those sheets. The snapshot contains 1,464 stored cells, including 1,364 formula cells with saved caches.

The initial exploration target is `Sheet9.CmdPremium_Click`:

1. Overview identifies `PremiumTable` and its declared button/event association.
2. Query the retained `Check Box 1` binding with owner sheet `CalculationOfBE_OtherDisease`. The delivered canonical record must show the resolved target `CalculationOfBE_OtherDisease!F14` and its original package-part sources. Query `F14` separately for the linked cell fact. A binding fact can pass source-citation validation; changing its target or citing an undelivered binding must fail.
3. Query the canonical ActiveX event record for `PremiumTable.CmdPremium`, followed by the complete `Sheet9` module. This checks the control-to-module evidence path without claiming that the button was executed.
4. Query `ValueTable`, `Variables` and `IssueAge` declarations with the appropriate worksheet context. Query the cells holding the variable names and their declared destinations. Treat a current table entry as an observed value, not proof of every possible dynamic target.
5. Query `GP`, `NLP` and the other output names read by the procedure. Trace the corresponding formula paths into retained inputs and preserved dependency ranges.
6. Record the source sequence: write Main inputs, loop over issue age, call `Application.Calculate`, read Premium outputs and write PremiumTable results. This is a static interpretation until observed at runtime.
7. Validate source citations and independently review the resulting workflow specification.

If changed-scenario behavior is required, capture a baseline, one input variation and the relevant macro operation in an Excel copy. Record inputs, output values, calculation mode and the observed writes. These cases supply runtime evidence and later numerical reconciliation. They are not a prerequisite for preserving facts or building a usable index, and they do not silently broaden the four-sheet exclusion.

## Implementation order and checks

| Increment | Deliverable | Evidence that it works |
| --- | --- | --- |
| P0 Preparation and handoffs | Source-bound scope, separately addressable views with canonical control records, preparation manifest, paired English handoffs and prepare CLI | Existing final reports stay unchanged; control views are bound to the checked module artifacts; relocated roots resolve; scope changes invalidate only reading preparation; dry-run writes nothing |
| P1 Targeted querying | Query CLI, bounded packets, catalogue and agent continuation instructions | Reading one source/view does not load the monolithic views file; ambiguous names require context; pagination and oversized records are explicit |
| P2 Static dependency tracing and packet validation | Name/range expansion, bounded upstream/downstream trace, claim validation against canonical delivered evidence | GP tracing reaches its supported destination; consumers through ranges/names are found; cycles and limits terminate; altered or unsupplied facts fail validation |
| P3 Real scoped exploration | CmdPremium workflow specification, bounded typed judgments and fresh independent review | Every conclusion has evidence or an explicit unresolved decision; excluded controls stay outside the backlog; cached results are not presented as new-scenario results |

Before implementation, assess these increments separately. P0 and P1 have a clear path across existing CLI, schemas, views and index helpers and fit a normal implementation tier. P2 needs a separate assessment of name/range resolution and graph coupling; substantial resolution behavior can justify a complex tier. The initial target is static supported references, not arbitrary Excel evaluation. Missing design evidence triggers investigation rather than automatic tier escalation.

Required checks should concentrate on realistic workflow cases:

1. Prepare rejects wrong-source scope, missing native source references and checksum mismatch while preserving original artifacts. It derives dependency outputs from checked source data, ignores legacy sidecars as inputs, and detects a changed or missing generated snapshot/audit on reuse or validation.
2. An unchanged reading bundle resumes; changed scope or source/compiler inputs invalidate the relevant reuse.
3. Query packets preserve complete source facts and locations and return only the selected eligible source records. The retained OtherDisease checkbox and PremiumTable ActiveX event have canonical view/record citations bound to their module artifacts.
4. Workbook-local and worksheet-local names, identical names across batch sources, empty cells and oversized VBA records follow the declared query rules.
5. Tracing resolves supported absolute names/ranges, reports unsupported/dynamic boundaries, handles reverse range membership and stops at its limits.
6. The 12 approved Pricing ranges remain readable dependencies without converting their excluded parent sheets or restoring the 265 ignored repairs. A contained query to `CalculationOfBE_CI!O7:Q112` succeeds; a direct query to `CalculationOfBE_CI!A1` or a wider range crossing outside the approved range returns `out_of_scope`. An excluded sheet overview returns metadata only, and tracing stops at the approved dependency data boundary.
7. Claim validation rejects an altered fact, wrong run/location, tampered canonical reference or citation to a record not supplied in the packet.
8. Discovery and handoff commands return structured errors and non-zero exits for failed integrity or validation; reading operations do not exhaust recovery history.

The final implementation review should verify both the fixed progression and the amount of repeated work. Instrument raw extraction, inventory parsing, graph construction and selected view loading in the real walkthrough. Preparation may perform the expensive work once; repeated questions should reuse it. Performance claims require observed counts and timings from that walkthrough.
