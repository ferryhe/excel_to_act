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
| B1 | `docs/design/layered_architecture.md` | Not created |
| B2 | `docs/design/agent_reading_contract.md` | Not created |
| B3 | `docs/design/fidelity_rules.md` | Not created |
| C1 | `docs/design/views_l1_compiler.md` | Not created |
| E1/E2 | `docs/experiments/*` | Not created |
| F1/F2 | `docs/adr/*` | Not created |

---

## 0. Current State Alignment (As-is)

This is a selection review for a **project already in progress**, not pre-project research. Current repository capabilities:

| Module | File | Status | Relevance to this document's conclusions |
|---|---|---|---|
| Manifest reader | `src/excel_to_act/ingest/openpyxl_reader.py` | Implemented | Confirmed in §4.6 |
| OOXML package scanning | `src/excel_to_act/ingest/ooxml_package.py` | Implemented | §4.6b (omitted from v1 review; now added) |
| Cell inventory | `src/excel_to_act/inventory/extractor.py` | Implemented | Confirmed in §4.6; dual-view load reads formulas and cached values, and records a warning when a cache is missing |
| Dependency graph | `src/excel_to_act/graph/builder.py` | Implemented (regex parsing) | §6.5 concludes it needs replacement |
| Classification/confirmation/storage/orchestration/CLI | `classify/`, `confirm/`, `store/local_store.py`, `orchestrator/phase1.py`, `interfaces/cli.py` | Implemented | Outside this document's scope (belongs to B1) |
| Report | `src/excel_to_act/report/handoff.py` | Step 1 JSON/Markdown handoff implemented | General report generation is outside this document's scope (belongs to B1 / D2) |
| Plugin protocols | `src/excel_to_act/plugins/contracts.py` | 6 Protocols, **no OracleRunner** | Add one for L3 in §7 |
| Optional dependencies | `pyproject.toml [project.optional-dependencies].formula` | `formulas>=1.3` + `xlcalculator>=0.5` bundled together | §7 recommends splitting them |

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
| **Formal definition** of the layered architecture and inter-layer contracts | **B1** (this document's §7 is only a **non-normative draft** to inform B1) |
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
| **License** | Project convention: **avoid GPL/AGPL-style copyleft in the core dependency path** |
| **Activity** | Date of the most recent release |
| **Suggested layer** | Draft pending B1 confirmation (L0 ingestion / L3 validation / reference only) |

---

## 3. Summary of Conclusions (TL;DR)

1. **No existing library can serve as the core.** The current options are all **calculation engines** (answering "what does it calculate?"); this project needs a **lossless extractor** (answering "what is in the workbook, where is it, and in what format?"). An engine discards source information; the extractor must preserve it.
2. **The strongest option, `formulas`, is also the riskiest dependency:** it has the highest coverage (90.1%) and is the most active, but its license is **EUPL 1.1+ (copyleft)**.
3. **`pycel` is ruled out:** GPLv3 + inactive since 2021.
4. **The in-house core ingestion layer (openpyxl) is implemented.** This document confirms that choice and identifies optional enhancements (stdlib package scanning, replacing regex with a tokenizer) and a validation path (L3 oracle).
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
| EUPL 1.1+ | formulas | **No** | Copyleft; lower risk for internal use, legal review required for distribution with artifacts |
| GPLv3 | pycel, Gridmonger | **No** | Conflicts with project convention |
| MPL 2.0 | LibreOffice, orcus/ixion | N/A | Out-of-process invocation or reference only; not linked |
| BSD-3 | xlwings | No | Depends on Excel COM |

> **Stricter interpretation:** The plan says "Avoid **GPL/AGPL** core dependencies." This document also classifies **EUPL 1.1+** as prohibited from core dependencies; this is a **stricter interpretation** that F1 ADR must confirm and that must be reflected in the corresponding section of `docs/plans/phase1_excel_decomposition_plan.md`.
> **Transitive dependencies:** The table covers direct dependencies only. Transitive dependencies such as `schedula` for `formulas` and numpy/openpyxl for `xlcalculator` **must also be verified** (see §8).

---

## 6.5 Formula Reference Parsing (Tokenizer / AST) Focus

The plan explicitly names `formulas` / `xlcalculator` for **formula reference parsing / evaluation** experiments. This document answers each use case separately:

| Use | Conclusion |
|---|---|
| **Reference parsing** | Prefer **`openpyxl.formula.tokenizer`** (already a core dependency, MIT, **no new dependency and no license risk**). **Do not add EUPL/GPL dependencies for reference parsing.** |
| **Evaluation** | Use `formulas` only as an isolated L3 oracle (EUPL); refer to pycel tokenizer ideas only (GPLv3, **do not import code**). |

The current implementation uses regex (`graph/builder.py:20`) and has three problems: false address matches inside string literals, failure to recognize `Table[Col]`, and no defined-name nodes.

---

## 7. Placement Recommendations (**Non-normative Draft, Pending B1 Confirmation**)

| Layer | Choice | Existing modules | Missing modules |
|---|---|---|---|
| **L0 Ingestion** | `openpyxl` + stdlib `zipfile`/XML (core); `fastexcel` for side-path acceleration only | `ingest/openpyxl_reader.py`, `ingest/ooxml_package.py`, `inventory/extractor.py` | `inventory/layout.py`, `inventory/opaque.py`, `ingest/calamine_reader.py` |
| **L1 View** | **In-house; no ready-made option** | None (only `SourceLocation` in `schemas/artifacts.py`) | Recommend adding `view/` + `ViewSlice` contract |
| **L2 Semantic reasoning** | In-house | `classify/rules.py`, `classify/classifier.py`, `confirm/templates.py` | No actual semantic reasoning module; `ActuarialHint` (6 values) differs from the plan's 10 hints |
| **L3 Validation oracle** | `formulas` (optional dependency, isolated) + LibreOffice (out of process) + **cached-value comparison (preferred, no dependency)** | Formula cached-value collection implemented; numerical reconciliation and reporting still missing | `validation/{formulas_oracle,libreoffice_oracle}.py` + `OracleRunner` protocol + `ValidationReport` |
| Outside layers (B1 must address) | — | `store/local_store.py`, `orchestrator/phase1.py`, `interfaces/cli.py` implemented | `report/markdown.py` missing (`jinja2` is declared but unused) |

**⚠ Prerequisite for cached-value comparison:** Dual-view loading with `data_only=True` and the `cached_value` field are implemented; if a workbook has not been recalculated by Excel, a usable cache may still be missing, and a warning will be recorded. Cached-value collection is available, but the L3 reconciliation workflow is not implemented yet.

**⚠ Current extras:** `pyproject.toml` declares `formula = ["formulas>=1.3", "xlcalculator>=0.5"]`, bundling an EUPL package and an MIT package in the same extra, contrary to this document's isolation recommendation. Recommend splitting into `oracle-formulas` (EUPL, explicit opt-in) and `xlcalc` (MIT).

---

## 8. Items to Verify (Open Questions)

1. Exact license and formula/format-reading capabilities of `fastexcel` / `python-calamine` → test before adoption.
2. Boundary under EUPL between "internal use vs. distribution with artifacts" for `formulas` → legal confirmation needed (→ F1).
3. Failure modes of `formulas` on workbooks containing VBA / pivot tables → test with fixtures.
4. **Python version feasibility:** `requires-python = ">=3.11"` has no upper bound; the CI core dependency matrix tests 3.11/3.12/3.13; optional dependencies are not yet in that install matrix. Test whether `pip install formulas / xlcalculator / fastexcel` installs and imports successfully.
5. Whether the built-in **openpyxl tokenizer** (`openpyxl.formula.tokenizer`, openpyxl≥3.1) covers A1 / cross-sheet / structured references / defined names → if so, prefer it for reference parsing and replace regex.
6. **Transitive dependency licenses:** `formulas` → `schedula`, etc.; `xlcalculator` → numpy, etc.
7. How the oracle degrades to a warning rather than an error when **cached values are missing** (workbook has never been recalculated and saved by Excel).
8. License and API stability of `msoffcrypto-tool`; formula/format coverage of `.xlsb` by `pyxlsb`.
9. Which fidelity dimensions `openpyxl`'s `read_only=True` mode specifically loses → test.

---

## 9. Acceptance Self-Check

| Acceptance item | Status | Evidence | Blocker |
|---|---|---|---|
| Each tool (§4.1–§4.8) has a three-line conclusion | **Complete** | Each section 4.1–4.8b has a three-line conclusion block | — |
| Main calculation-engine candidates are covered | **Complete** (13 tools / 11 sections) | Entire §4; §4.9 contains fallback references | Not included: SheetJS, pandas, xlrd (none is positioned as "Excel → code/evaluation") |
| No overlap with A2/A3/B1/B2/B3/C1 | **Partially complete** | §1 Out of scope states the non-normative boundary | Six related documents have not been created; **review after their creation** |
| Aligned with repository state (references `src/` files) | **Complete** | §0 current-state table, §4.6 gap list, §7 existing/missing columns | — |
| Placement recommendations map to specific files | **Complete** | §7 table's "Existing modules / Missing modules" columns | — |
| Converted into actionable work (PR / extras adjustment) | **Not complete** | — | Add PR-13 to `docs/plans/pr_plan_phase1.md` (L3 oracle + `OracleRunner`, reusing implemented cached-value collection) and adjust `pyproject.toml` extras |
| Version/license data is verifiable | **Partially complete** | PyPI metadata pages + GitHub README | §8.4/§8.6 need testing; `FlyingKoala` lacks an exact release date |

---

## Appendix: Code-Side Gaps (Outside This Document's Decisions; Refer to B3 / C-Class Documents / Separate Issues)

1. The coverage invariant in `inventory/extractor.py:100` is an **identity** (`discovered = recognized + opaque`), so it is always true and cannot detect silent data loss; `discovered` must be enumerated independently with `zipfile`.
2. `CellInventory.cached_value` and `cached_value_available` are implemented and match `inventory/README.md`; the cached-value reconciliation oracle is still unimplemented.
3. The `ActuarialHint` enum (6 values) differs from the 10 hints on lines 88–97 of the plan.
4. `zipfile.ZipFile` at `ooxml_package.py:51` has no exception handling; encrypted/damaged files raise an uncaught exception.
5. `tests/` is not split by module (only 2 files), and `examples/` has no static fixture.
