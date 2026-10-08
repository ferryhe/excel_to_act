# A1 · Excel Tooling Survey: Calculation Engines and "Excel → Code" Projects

- **Status:** Draft v2 (independent review feedback incorporated; awaiting final confirmation)
- **Data verified on:** 2026-10-02
- **Data sources:** Each project's PyPI metadata page, GitHub README, and source code
- **Document type:** Research
- **Review:** An independent subagent cross-checked the full document against `docs/plans/*`, `pyproject.toml`, and `src/**`.

## Related Documents (Path + Current Status)

| ID | Path | Status |
|---|---|---|
| A2 | `docs/research/docling_fidelity_assessment.md` | Not created |
| A3 | `docs/research/llm_table_encoding.md` | Not created |
| B1 | `docs/design/layered_architecture.md` | Implemented; normative current-state map |
| B2 | `docs/design/agent_reading_contract.md` | Implemented |
| B3 | `docs/design/fidelity_rules.md` | Implemented |
| C1 | `docs/design/views_l1_compiler.md` | Implemented |
| E1/E2 | `docs/experiments/*` | Not created |
| F1/F2 | `docs/adr/*` | Not created |

---

## 0. Current State Alignment (As-is)

This is a selection review for a **project already in progress**, not pre-project research. Current repository capabilities:

| Module | File | Status | Relevance to this document's conclusions |
|---|---|---|---|
| Manifest reader | `src/excel_to_act/ingest/openpyxl_reader.py` | Implemented | Confirmed in §4.6 |
| OOXML package scanning | `src/excel_to_act/ingest/ooxml_package.py` | Implemented | §4.6b (omitted from v1 review; now added) |
| Cell inventory and exact source facts | `src/excel_to_act/inventory/extractor.py`, `src/excel_to_act/steps/step1/source_scan.py`, `src/excel_to_act/steps/step1/workflow.py` | Implemented | Normalized inventory and separately retained raw facts/cache fields; see #13 in the architecture map |
| Dependency graph | `src/excel_to_act/graph/builder.py` | Implemented (regex parsing) | §6.5 concludes it needs replacement |
| Deterministic views and evidence navigation | `src/excel_to_act/views/__init__.py`, `src/excel_to_act/steps/step2/` | Implemented | Workbook/sheet/region views and source-bound evidence; see #14 in the architecture map |
| Rule classification and confirmation prompts | `src/excel_to_act/classify/`, `src/excel_to_act/confirm/` | Implemented in legacy `inspect`; optional Step 1 tools exist | Rule outputs and human questions, not README Step 3's broader semantic analysis |
| Storage/orchestration/interfaces/reporting | `src/excel_to_act/store/local_store.py`, `src/excel_to_act/orchestrator/phase1.py`, `src/excel_to_act/interfaces/cli.py`, `src/excel_to_act/report/handoff.py` | Implemented | Cross-layer services; report output spans stages (see B1) |
| Plugin protocols | `src/excel_to_act/plugins/contracts.py` | 6 Protocols, **no OracleRunner** | Add one for L3 in §7 |
| Optional dependencies | `pyproject.toml [project.optional-dependencies]` | `formulas>=1.3` is isolated in `oracle-formulas`; `xlcalculator>=0.5` is in `xlcalc` | §7 records the research concern; see the Issue #9 decision record for the default dependency gate |

---

## 1. Scope

### In scope

Evaluate tools that **"understand Excel formulas and calculate them / compile them to code"** (calculation engines / compilers / evaluators), plus tools related to **input formats**, and answer:

> Is there an existing open-source library that can serve directly as a **core** dependency for excel_to_act? If not, where can each option fit in the architecture?

### Out of scope

| Out of scope | Owner |
|---|---|
| Fidelity of "document parsers" such as Docling / markitdown / pandas rendering | **A2** |
| Encoding / compression / sampling strategies for LLMs | **A3** |
| **Formal definition** of the layered architecture and inter-layer contracts | **B1** (see [layered architecture](../design/layered_architecture.md); §7 here is a placement summary only) |
| Final decision record on whether to adopt | **F1** |
| Integration code and PR split for **new** tools | C-class module design / `docs/plans/pr_plan_phase1.md` (use `src/` as the authority for implemented behavior) |
| Performance benchmarks | E3 (not created) |
| Test strategy and fixtures | `tests/README.md` + D-class implementation plan |
| CI implementation of license compliance | F1 ADR + CI workflow |
| Python versions and runtime environment matrix | `pyproject.toml` + `.github/workflows/ci.yml` |

---

## 2. Evaluation Criteria

| Criterion | Why it matters |
|---|---|
| **Core capability** | Is it for "evaluation," "generating readable code," or "reading only"? |
| **Function coverage** | Availability of LOOKUP / INDEX / MATCH / OFFSET / financial functions commonly used in actuarial models |
| **Artifact form** | Does it produce an auditable, versionable intermediate representation? |
| **Structural coverage** | Availability of defined names / structured references / volatile / circular / styles |
| **License** | Project policy: **exclude GPL, AGPL, and EUPL license families from the default runtime dependency closure** |
| **Activity** | Date of the most recent release |
| **Suggested layer** | Placement recommendations are aligned with normative B1 (L0 ingestion / L3 validation / reference only) |

---

## 3. Summary of Conclusions (TL;DR)

1. **No existing library can serve as the core.** The current options are all **calculation engines** (answering "what does it calculate?"); this project needs a **lossless extractor** (answering "what is in the workbook, where is it, and in what format?"). An engine discards source information; the extractor must preserve it.
2. **The strongest option, `formulas`, is also the riskiest dependency:** it has the highest coverage (90.1%) and is the most active, but its license is **EUPL 1.1+ (copyleft)**.
3. **`pycel` is ruled out:** GPLv3 + inactive since 2021.
4. **The in-house core ingestion layer is implemented.** `openpyxl` reads workbook structures, and the Python standard library scans OOXML package parts. Future improvements include replacing regex formula-reference parsing with a tokenizer; L3 numerical verification remains a separate planned path.
5. **The current implementation does not yet realize openpyxl's full fidelity potential** (see the gap list in §4.6) — that is the real bottleneck to achieving "fidelity," not library selection.

---

## 4. Individual Evaluations

> Rule: Each tool in §4.1–§4.8 must have a **three-line conclusion** (layer / license / latest release); §4.9 contains fallback references and is exempt from this rule.

### 4.1 `formulas` (vinci1it2000)

- **Position:** Excel formula interpreter + workbook compiler
- **Version / activity:** 1.3.4, 2026-03-11 (active)
- **License:** **EUPL 1.1+** (copyleft, compatible with GPL)
- **Capabilities:** `ExcelModel().loads().finish().calculate()` compiles an entire workbook into a Python execution graph (schedula dispatch); supports circular references with `circular=True`; extracts submodels with `from_ranges()`; supports `compile(inputs, outputs)`; models can be exported as JSON; CLI `build/calc/test/serve`, with batch multi-scenario support
- **Function coverage:** **483 / 536 = 90.1%** (DATE&TIME / ENGINEERING / FINANCIAL / STATISTICAL / LOGICAL / OPERATORS are all 100%; LOOKUP 82.5%, MATH&TRIG 88.8%, TEXT 88.0%; **CUBE / DATABASE / WEB / AUTOMATION are 0%**)
- **Limitations:** Output is not readable Python source code (it is an internal dispatch graph); it does not produce a coordinate-level fidelity inventory; CUBE/DATABASE/WEB are missing

> 1. **Layer:** L3 validation oracle (optional dependency, isolated process/optional extra)
> 2. **Allowed in core dependencies:** **No** (EUPL 1.1+ copyleft)
> 3. **Latest release:** 2026-03-11

### 4.2 `pycel`

- **Version:** 1.0b30, **2021-10-13** (inactive for 4+ years)
- **License:** **GPLv3**
- **Capabilities:** Generates graph-based Python (caching + lazy evaluation); supports array CSE, INDIRECT/OFFSET/INDEX, structured references, and iterative calculation
- **Limitations:** GPLv3; inactive; does not compile VBA; OFFSET depends on whether cells have been compiled

> 1. **Layer:** Research reference only (tokenizer / address-handling ideas, **do not import code**)
> 2. **Allowed in core dependencies:** **No** (GPLv3)
> 3. **Latest release:** 2021-10-13

### 4.3 `xlcalculator` (modernized version of koala2)

- **Version:** 0.5.0, **2023-02-06**; **License: MIT** (changed from GPL-3 to MIT starting in 0.2.3)
- **Capabilities:** Workbook → Python state (can be saved/loaded); focuses on submodels; evaluates individual cells/named ranges/ranges/shared formulas; custom functions can be registered
- **Limitations:** **Refactored to AST evaluation starting in 0.1.0; it no longer generates Python code**; does not support array/CSE formulas; **missing INDEX / OFFSET / INDIRECT / HLOOKUP / COLUMN / ROW**; VLOOKUP supports exact match only; LN / YEARFRAC differ from Excel
- **Difference from the plan:** The plan allows it as an in-plugin experiment; due to the capability gaps above, this document **does not recommend spending engineering time on it**

> 1. **Layer:** Research reference (AST design and precision-handling sections)
> 2. **Allowed in core dependencies:** Allowed (MIT), but not recommended because capabilities are insufficient
> 3. **Latest release:** 2023-02-06

### 4.4 `koala2`

Predecessor to `xlcalculator`, deprecated. Do not use.

> 1. **Layer:** Do not use
> 2. **Allowed in core dependencies:** Not applicable (deprecated)
> 3. **Latest release:** Maintenance discontinued

### 4.5 `FlyingKoala`

- **Position:** A set of xlwings helper functions that let Python call and unit-test Excel formulas; depends on xlwings + koala + pandas
- **Value:** Methodology for "using Excel itself as the calculation engine to unit-test formulas" and for incremental migration
- **Limitations:** Depends on Excel COM, **so it cannot run in CI (ubuntu-only)**

> 1. **Layer:** L3 validation reference (regression methodology), **not in CI**
> 2. **Allowed in core dependencies:** No (depends on Excel COM)
> 3. **Latest release:** N/A (no recent release; repository has low activity, exact date needs manual confirmation)

### 4.6 `openpyxl` (**L0 core**)

- **License:** MIT; **version:** `pyproject.toml` pins `openpyxl>=3.1`
- **Decisive capabilities:**
  - `data_only=False` → **original formula**; `data_only=True` → **cached value**
  - `number_format`, `data_type`, `comment`, `hyperlink`, merged cells, tables, defined names, data validation, conditional formatting, sheet protection/visibility
- **Limitations:** Does not calculate; slow on large files (`read_only=True` can speed up streaming, but loses some formatting/comment information, trading off against the fidelity goal)
- **⚠ The current implementation does not yet realize its full fidelity potential (gap list):**
  1. Cached values are now collected by a second load with `data_only=True` in `ingest/cached_values.py`; if the workbook has not been recalculated, a cache may still be unavailable and a warning is recorded.
  2. Reference parsing uses regex (`graph/builder.py:20` `REF_RE`), does not recognize structured references such as `Table[Col]`, and does not construct defined-name nodes.
  3. Fills / fonts / borders / print areas are not collected (required by plan bucket 6).
  4. Volatile / circular items are not identified (required by plan bucket 5).

> 1. **Layer:** **L0 core ingestion layer (implemented)**
> 2. **Allowed in core dependencies:** **Yes** (MIT)
> 3. **Latest release:** Actively maintained

### 4.6b Python standard library `zipfile` + `xml.etree` (OOXML package scanning)

- **License:** PSF / PSF-2.0; **no new dependencies**
- **Purpose:** Detect workbook parts that openpyxl cannot model: `xl/vbaProject.bin`, `charts/`, `pivotTables/`, `externalLinks/`, `connections.xml`, `drawings/`, `embeddings/`, and `media/`, then emit `UnsupportedFeature` — **this is the only implementation mechanism for the project's "do not silently drop" convention**
- **Implemented:** `src/excel_to_act/ingest/ooxml_package.py` (`OPAQUE_MARKERS`)
- **⚠ Known defect:** `zipfile.ZipFile(...)` at `ooxml_package.py:51` has no exception handling, so encrypted/damaged workbooks raise an uncaught exception instead of producing `UnsupportedFeature`.

> 1. **Layer:** **L0 core ingestion layer (implemented alongside openpyxl)**
> 2. **Allowed in core dependencies:** **Yes** (standard library)
> 3. **Latest release:** Released with CPython

### 4.7 `fastexcel` / `python-calamine`

- **Position:** Rust bindings for calamine, fast and read-only; **license:** MIT family (verify before adoption)
- **⚠ Constraint (must be followed):** Use only for **side-path acceleration / pre-screening**, **never as the sole L0 data source**. It must implement the `WorkbookReader` protocol, label artifacts `fidelity='values_only'`, and produce `unsupported_or_opaque` records for uncovered dimensions; otherwise it directly conflicts with the coverage invariant.

> 1. **Layer:** Optional L0 side-path acceleration (not a data source)
> 2. **Allowed in core dependencies:** To be verified
> 3. **Latest release:** Actively maintained

### 4.8 LibreOffice（headless / soffice）

- **License:** MPL 2.0; **usage:** invoke recalculation and conversion out of process; do not link
- **Value:** A second L3 oracle; has both value and limitations (requires installation, starts slowly, and has known numerical differences from Excel)

> 1. **Layer:** L3 validation oracle (out of process)
> 2. **Allowed in core dependencies:** Not applicable (out-of-process invocation, not linked)
> 3. **Latest release:** Actively maintained

### 4.8b `xlwings` (oracle runner named in the plan)

- **License:** BSD-3-Clause
- **Value:** Recalculates with Excel itself as the numerical baseline
- **Limitations:** Depends on Excel COM; CI uses `ubuntu-latest` and Python 3.11/3.12/3.13 (`.github/workflows/ci.yml:19`), so it **can run only on a developer's local Windows machine, not in CI**.

> 1. **Layer:** L3 oracle (local baseline, not in CI)
> 2. **Allowed in core dependencies:** No
> 3. **Latest release:** Active (exact version to be verified)

### 4.9 Other Research References (Exempt from the Three-Line Conclusion Rule)

| Project | Position | License | Conclusion |
|---|---|---|---|
| `Gridmonger` | JVM/Kotlin Excel reverse-engineering visualization | GPLv3 | Algorithm / interaction reference only |
| `orcus` / `ixion` | C++ spreadsheet model import and formula engine | MPL 2.0 | Syntax / engine reference only |
| `PyXLL` / `Excel-DNA` | Python embedded in Excel | Commercial / mixed | Opposite direction; not applicable |
| `SpreadsheetConverter` | Excel → C#/Java | Commercial | Do not adopt |
| `Mito` | Records in Jupyter to generate pandas code | Open source | Records operations rather than parsing existing models; not applicable |

### 4.10 Input Formats: Encrypted / Protected / Binary and Legacy Formats

| Format | Candidate tool | Notes |
|---|---|---|
| Encrypted / protected workbook | `msoffcrypto-tool` | License and API stability to be verified |
| `.xlsb` | `pyxlsb` | Formula/format coverage to be verified |
| `.xls` / legacy format | LibreOffice conversion | Convert out of process, then use the standard workflow |

**Hard constraint (project convention):** Any read failure must produce `UnsupportedFeature(severity=error)` and **must not raise an uncaught exception**. `ooxml_package.py:51` currently violates this constraint.

---

## 5. Comparative Matrix

| Tool | Produces readable code | Function coverage | License | Latest release | Evaluates formulas | Structural fidelity |
|---|---|---|---|---|---|---|
| `formulas` | No (dispatch graph + JSON) | **90.1%** | EUPL 1.1+ | 2026-03 | Yes | No |
| `pycel` | Yes (graph-based Python) | Not measured (depends on requirements) | GPLv3 | 2021-10 | Yes | No |
| `xlcalculator` | No (AST evaluation) | Medium-low | MIT | 2023-02 | Yes | No |
| `koala2` | — | Low | — | Deprecated | Yes | No |
| `openpyxl` | N/A | N/A | **MIT** | Active | No (**but can read Excel cached values**) | **Yes** |
| stdlib `zipfile`+XML | N/A | N/A | PSF | With CPython | No | **Yes (at part level)** |
| `fastexcel` | N/A | N/A | MIT (to be verified) | Active | No | Partial (values-only) |
| LibreOffice | N/A | High | MPL 2.0 (out of process) | Active | Yes | Partial |

**Key diagonal:** `openpyxl` is the only option with **fidelity potential**, yet it does not calculate. Whether that potential is realized depends on the completeness of our implementation (see the gap list in §4.6).

---

## 6. License Risk Summary

| License | Tool | Allowed in core | Notes |
|---|---|---|---|
| MIT | openpyxl, xlcalculator, fastexcel | **Allowed** | No copyleft propagation |
| PSF | stdlib zipfile/xml | **Allowed** | Standard library |
| EUPL 1.1+ | formulas | **No** | Excluded from the default runtime closure by project policy |
| GPLv3 | pycel, Gridmonger | **No** | Excluded from the default runtime closure by project policy |
| MPL 2.0 | LibreOffice, orcus/ixion | N/A | Out-of-process invocation or reference only; not linked |
| BSD-3 | xlwings | No | Depends on Excel COM |

> **Project dependency-selection policy:** The default runtime closure excludes GPL, AGPL, and EUPL license families. This is a project policy, not legal advice or a legal conclusion. CI checks installed package metadata for the default runtime closure; see the Issue #9 decision record for metadata sources and mixed-license handling.
> **Optional extras:** The default runtime gate excludes opted-in packages. Their transitive dependencies require separate review before adoption; the `oracle-formulas` extra is exercised in CI for calculation compatibility but remains outside the default closure.

---

## 6.5 Formula Reference Parsing (Tokenizer / AST) Focus

The plan explicitly names `formulas` / `xlcalculator` for **formula reference parsing / evaluation** experiments. This document answers each use case separately:

| Use | Conclusion |
|---|---|
| **Reference parsing** | Prefer **`openpyxl.formula.tokenizer`** (already a core dependency, MIT, **no new dependency and no license risk**). **Do not add EUPL/GPL dependencies for reference parsing.** |
| **Evaluation** | Use `formulas` only as an isolated L3 oracle (EUPL); refer to pycel tokenizer ideas only (GPLv3, **do not import code**). |

The current implementation uses regex (`graph/builder.py:20`) and has three problems: false address matches inside string literals, failure to recognize `Table[Col]`, and no defined-name nodes.

---

## 7. Placement Recommendations (summary; B1 is normative)

| Layer | Choice | Existing modules | Missing modules |
|---|---|---|---|
| **L0 Facts** | `openpyxl` + stdlib OOXML scanning | `ingest/`, `inventory/`, `steps/step1/source_scan.py`, `steps/step1/workflow.py` | #5/#13 remaining acceptance; no directory refactor implied |
| **L1 Deterministic views** | In-house, source-bound deterministic projection | `graph/builder.py`, `views/__init__.py`, `steps/step2/`, `schemas/step2_prepare.py`, `schemas/step2_query.py` | No new `view/` package or shared `ViewSlice` wrapper needed |
| **L2 Rule / semantic interpretation** | Current rules and human confirmation prompts | `classify/rules.py`, `classify/classifier.py`, `confirm/templates.py` | Broader README Step 3 semantic workflow remains a target; runtime enums are narrower than plan prose |
| **L3 Numerical verification** | Saved cache plus one actual recalculation source (#7) | Cache collection in `ingest/cached_values.py` | `validation/`, `OracleRunner`, and `ValidationReport` do not exist; generated-code equivalence is a separate later check |
| Cross-layer services | Persist, coordinate, expose interfaces, share contracts, render reports | `store/local_store.py`, `orchestrator/phase1.py`, `interfaces/cli.py`, `schemas/`, `plugins/`, `report/handoff.py` | Reports can span phases; do not force them into one layer |

**⚠ L3 boundary:** Cached-value collection is implemented, but the numerical comparison workflow is not. #7 requires a saved workbook cache and a distinct actual recalculation result; copying a cache is not an oracle. This pre-generation workbook-baseline check is not the later post-generation code-equivalence comparison.

**Default dependency policy:** `formulas` is excluded from the default install and remains available only through the explicit `oracle-formulas` extra; `xlcalculator` remains separate in `xlcalc`. CI checks the installed default runtime packages and fails closed on missing or unclassifiable license metadata. The Issue #9 decision record documents the mixed-license rule and metadata sources.

**Issue #10 deferral:** There is no concrete Docling/markitdown integration use case now, and the existing `openpyxl` + OOXML path has direct work items for known gaps. The focused parser-fidelity comparison is incomplete; latest third-party versions have not been reassessed, so this survey does not show that their fidelity is disproven. Reopen only for a concrete workbook use case current ingestion/views cannot serve, then pin versions and compare fixtures against that use case's fidelity requirements. No new parser research was performed for this update.

---

## 8. Items to Verify (Open Questions)

1. Exact license and formula/format-reading capabilities of `fastexcel` / `python-calamine` → test before adoption.
2. If a future distribution bundles the opt-in `formulas` backend, assess that specific decision separately; the current dependency policy makes no legal conclusion.
3. Failure modes of `formulas` on workbooks containing VBA / pivot tables → test with fixtures.
4. **Python version feasibility:** `requires-python = ">=3.11"` has no upper bound; CI tests the default runtime and `oracle-formulas` fixture on Python 3.11/3.12/3.13. Other optional dependencies and future Python versions have not been tested.
5. Whether the built-in **openpyxl tokenizer** (`openpyxl.formula.tokenizer`, openpyxl≥3.1) covers A1 / cross-sheet / structured references / defined names → if so, prefer it for reference parsing and replace regex.
6. **Optional dependency licenses:** the default runtime closure is gated in CI; transitive licenses for unselected optional extras such as `xlcalculator` remain unverified before adoption.
7. How the oracle degrades to a warning rather than an error when **cached values are missing** (workbook has never been recalculated and saved by Excel).
8. License and API stability of `msoffcrypto-tool`; formula/format coverage of `.xlsb` by `pyxlsb`.
9. Which fidelity dimensions `openpyxl`'s `read_only=True` mode specifically loses → test.

---

## 9. Acceptance Self-Check

| Acceptance item | Status | Evidence | Blocker |
|---|---|---|---|
| Each tool (§4.1–§4.8) has a three-line conclusion | **Complete** | Each section 4.1–4.8b has a three-line conclusion block | — |
| Main calculation-engine candidates are covered | **Complete** (13 tools / 11 sections) | Entire §4; §4.9 contains fallback references | Not included: SheetJS, pandas, xlrd (none is positioned as "Excel → code/evaluation") |
| Cross-document scope/overlap review | **Partially complete** | B1, B2, B3, and C1 exist and are listed above | A2/A3 remain uncreated; review cross-document scope after they are available |
| Aligned with repository state (references `src/` files) | **Complete** | §0 current-state table, §4.6 gap list, §7 existing/missing columns | — |
| Placement recommendations map to specific files | **Complete** | §7 table's "Existing modules / Missing modules" columns | — |
| Converted into actionable work (PR / extras adjustment) | **Partially complete** | Issue #9 adds the default-runtime policy gate and validates `formulas` as the selected oracle | Add PR-13 to `docs/plans/pr_plan_phase1.md` for L3 oracle integration and `OracleRunner`; no extras adjustment is needed |
| Version/license data is verifiable | **Partially complete** | CI declares Python 3.11/3.12/3.13; local checks on Python 3.11 passed | Cross-version CI results and optional-extra licenses beyond `oracle-formulas` remain to be verified; `FlyingKoala` lacks an exact release date |

---

## Appendix: Code-Side Gaps (Outside This Document's Decisions; Refer to B3 / C-Class Documents / Separate Issues)

1. The coverage invariant in `inventory/extractor.py:100` is an **identity** (`discovered = recognized + opaque`), so it is always true and cannot detect silent data loss; `discovered` must be enumerated independently with `zipfile`.
2. `CellInventory.cached_value` and `cached_value_available` are implemented and match `inventory/README.md`; the cached-value reconciliation oracle is still unimplemented.
3. The `ActuarialHint` enum (6 values) differs from the 10 hints on lines 88–97 of the plan.
4. `zipfile.ZipFile` at `ooxml_package.py:51` has no exception handling; encrypted/damaged files raise an uncaught exception.
5. `tests/` is not split by module (only 2 files), and `examples/` has no static fixture.
