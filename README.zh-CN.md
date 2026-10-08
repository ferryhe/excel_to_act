# excel_to_act

[English](README.md) | 简体中文

项目 README 提供英文和简体中文版；其他文档、代码注释和 docstring 统一使用英文。

Excel-to-actuarial-model tooling：**CLI + API + agent-ready**，可按 skill 安装给任意 agent 使用。

第一原则：**先做无损分解（lossless decomposition），再做理解与生成**。任何读不懂的东西必须记为 opaque，禁止静默丢弃。

```text
recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects
```

> ⚠️ 范围说明：`CoverageSummary.discovered_workbook_objects` 来自独立 OOXML 源扫描，recognized 数量也按匹配到的源身份统计；仅有汇总计数不能证明逐个身份都已覆盖。自 #5 / [PR #38](https://github.com/ferryhe/excel_to_act/pull/38) 起，旧 `inspect` 会运行 `verify_completeness`，比较源与 inventory 身份、校验存储计数，发现漏项时失败并以非零码退出。人工 Step 1 有独立的逻辑对象账本和 OOXML 部件账本。这些检查不表示所有 Excel 功能都已语义解析。见[Phase 1.5 待办与状态](docs/issues/backlog.md)及 §4 I。

---

## 1. 流水线总览

```mermaid
flowchart TD
    IN["input/ · 原始目录<br/>当前 Step 1: step1 convert &lt;dir&gt; --out<br/>legacy: inspect &lt;workbook&gt; --out"]
    S0["Step 0 · Intake &amp; Identify<br/>识别格式 / 加密 / 包完整性"]
    S1["Step 1 · Lossless Decomposition<br/>按内容类型全覆盖分类提取"]
    O1[("output/step1_decomposition/<br/>按种类 md + json")]
    V["verify_completeness<br/>独立重算对象全集并比对"]
    H["handoff.json + handoff.md<br/>Step 1 唯一出口"]
    S2["Step 2 · Artifact Index<br/>为后续 agent 建产物索引与摘要"]
    O2[("output/step2_index/<br/>INDEX.md + index.json")]
    S3["Step 3 · Semantic Analysis（目标）<br/>域分析 / 结构规约"]
    O3[("output/step3_analysis/")]
    B0["生成前工作簿基准检查（#7）<br/>已保存缓存 vs 实际重算"]
    S4["Step 4 · Structured Python Generation"]
    O4[("output/step4_generation/")]
    B1["生成后代码等价性对账<br/>生成代码 vs 工作簿基准"]

    IN --> S0 --> S1 --> O1 --> V --> H --> S2 --> O2 --> S3 --> O3 --> B0 --> S4 --> O4 --> B1

    style S1 fill:#dff5e1,stroke:#2e7d32,stroke-width:2px
    style O1 fill:#eef7ff,stroke:#1565c0
```

**落地状态**：旧版单工作簿 `inspect` 已生成依赖图、规则分类和供人工确认的问题，执行结构完整性检查及独立的缓存与可选 `formulas` 重算比较，并在 handoff 中分别记录结构与数值状态。当前人工 Step 1/2 流程转换并检查源事实、为 handoff 建索引、编译确定性视图，并支持证据查询和追踪。README Step 3 描述的是更广泛的语义分析目标；现有规则分类不等于完整语义分析。Step 2 批次 `index.json`/`INDEX.md` 与旧版 `artifact_index.json` 不同。**Python 生成和生成后等价性检查仍是后续工作**。当前契约和目标边界见[分层架构](docs/design/layered_architecture.md)。

上图展示完整目标路线，包含尚未完成的步骤；当前人工入口是 `excel-to-act step1 convert <raw-directory> --out <output-directory>`，旧 `inspect` 仍用于单工作簿 Phase 1 流程。

### 当前可用：人工 Step 1 → 人工 Step 2

新的 `step1` 命令接收原始目录，递归记录每个文件；`.xlsx` / `.xlsm` 进入转换，其它格式、损坏或无法读取的文件会保留为失败条目。每个输入按源 SHA-256、相对路径键和运行 ID 分开存放，批次 handoff 汇总全部输入，不使用“最后一个文件”的根目录别名。

Step 2 会为该 handoff 及其单源引用建立索引：

```text
excel-to-act step2 agent
excel-to-act step2 tools
excel-to-act step2 index --handoff PATH --step1-root DIR [--out DIR] [--resume]
excel-to-act step2 validate --index output/step2_index/index.json --step1-root DIR
output/step2_index/index.json
output/step2_index/INDEX.md
```

Step 2 批次索引与下文旧版单工作簿 `artifact_index.json` 不同。

若输出目录位于输入目录之内，目录扫描会排除该输出子树，并在批次 `discovery` 信息中记录排除路径；输入目录与输出目录完全相同时会明确报错。

```powershell
excel-to-act step1 tools                         # JSON 工具目录
excel-to-act step1 agent                         # 可复用的 host-agent 定义
excel-to-act step1 convert .\raw --out .\output # 第一步：事实提取与检查
excel-to-act step1 check --run <run-dir>         # 重新读取当前源文件并检查候选产物
excel-to-act step1 coverage --run <run-dir>      # 分别检查逻辑对象与 OOXML 部件
excel-to-act step1 fidelity --run <run-dir>      # 比较原始 XML 字段与标准化事实
excel-to-act step1 finalize --run <run-dir>      # 通过新检查后发布最终转换目录
```

也可以对单个工作簿分别检查复选框、ActiveX 和 VBA 交接。每类都有 `identify`、`convert`、`evaluate` 命令：`identify` 只读；`convert --dry-run` 只显示计划，不创建或改动输出；`evaluate` 会重新提取源事实并与已保存产物比较。

```powershell
excel-to-act step1 checkbox identify .\model.xlsm
excel-to-act step1 checkbox convert .\model.xlsm --out .\controls [--dry-run]
excel-to-act step1 checkbox evaluate .\model.xlsm --out .\controls
excel-to-act step1 activex identify .\model.xlsm
excel-to-act step1 activex convert .\model.xlsm --out .\controls [--dry-run]
excel-to-act step1 activex evaluate .\model.xlsm --out .\controls
excel-to-act step1 vba identify .\model.xlsm
excel-to-act step1 vba convert .\model.xlsm --out .\controls [--dry-run]
excel-to-act step1 vba evaluate .\model.xlsm --out .\controls
```

默认 Step 1 交接会写出类型化控件产物、供人阅读的 `controls_handoff.md`，以及带校验和索引的独立 VBA 源码文件。像源文件中的 `#REF!` 复选框链接会作为 invalid 记录，不会猜目标单元格。ActiveX 事件名按工作表 codeName 与已声明的 VBA 过程做静态关联；不会运行宏，ActiveX 二进制流仍保持 opaque。导出的 VBA 是可读源码交接，不是 Excel 运行时行为的转译。

默认门槛是 100% 逻辑对象可追溯、OOXML 部件逐字节保存、已支持事实精确一致。未知或未解析部件仍会保存；默认策略把它们报告为 `partial`，并只在其它检查通过后允许进入 Step 2。可以用 `--no-allow-opaque` 禁止这种交接。手动设置较低的 `--fidelity-min` 会明确产生 partial 结果；它不能绕过缺失的支持对象、损坏 XML、源文件变化或校验错误。Step 1 不计算公式、不判断精算含义；人确认 Step 1 handoff 后，Step 2 才开始建立索引并决定后续解释工作。`inspect <workbook> --out <dir>` 保持原来的单工作簿 Phase 1 行为。

包部件的解析比例和 opaque 比例使用同一单位（所有非目录 OOXML ZIP 部件），并与逻辑对象覆盖率分开：`parsed_coverage_ratio` 是已关联到 inventory 身份的支持逻辑对象比例，`parsed_package_parts_ratio` 是 fresh scan 中已读取且非 opaque 的部件比例，`opaque_rate` 是 opaque 部件比例。`--parsed-package-min` 和 `--opaque-max` 都接受 `[0,1]`，默认分别为 `0` 和 `1`，允许显式的 opaque `partial` 交接；提高前者或降低后者可以独立收紧包部件要求。`--no-allow-opaque` 无条件禁止 opaque 部件，任何数值门槛都不能覆盖缺失对象、源读取/校验错误或字节保存失败。

每个源文件的质量结果和工具响应都带有 `metrics_state`。`measured` 表示已成功 fresh scan；`unavailable` 表示当前源文件无法读取或校验，源派生的分母、计数、比例和保真差异会写成 JSON `null`，不会用旧报告或候选 ledger 代替当前测量。只有成功扫描后实测得到的零才写为数字 `0`；未能测量始终阻止放行，即使配置门槛为零也不能 finalize 或晋升。

`raw_value_text`、原始公式文本/属性、缓存文本和日期序列是源文件原文证据；下游若要求精确数值文本，应读取 `raw_value_text`。标准化数值会转换为整数或浮点数，长小数可能有浮点近似；日期按工作簿的 1900/1904 日期系统转换。公式缓存只表示文件中保存的结果，缺失缓存保持为缺失，转换过程不会重新计算公式。

这些扩展源事实只有 Step1 扫描实际测量后才有值；旧 `inspect` 输出或早期 JSON 中的 `null` 表示未测量，不表示 `false` 或 1900 日期系统。

字符串的标准化值拼接 OOXML 正文 `<t>` 与富文本 `<r><t>`，不把 `<rPh>` 注音或容器排版空白并入单元格文字；原始工作表和共享字符串部件仍按原字节保留并校验。

每批生成 `batches/<batch-id>/batch_handoff.{json,md}`，每个源文件有独立候选目录，其中包括 `source_facts.json`、`logical_objects.json`、`package_parts.json`、`quality.json` 和双格式 handoff。通过 `finalize` 的候选复制到唯一的 `final/` 目录；被阻断的候选仍有 handoff 和可操作的下一步建议。

**后续建造目标：Step 1 的无损分解扩展**（见 §4 的缺口清单）；目录式转换、独立覆盖/保真检查与双格式 handoff 已可通过上面的 `step1` 命令使用。

### 与 `phase1_excel_decomposition_plan.md` 的映射

| README 流水线 | plan 中的位置 |
|---|---|
| Step 0–3 | **plan Phase 1 范围内**（ingest → inventory → graph → classify → confirm → store → report） |
| Step 4（生成 Python） | plan 「Non-goals」明确排除 |
| 工作簿基准检查（#7） | 旧版 `inspect` 将已保存的公式缓存与可选 `formulas` 实际重算结果比较，并保存 `validation_report.json`；缺失覆盖会明确标记 |
| 生成代码等价性 | Step 4 之后对照已接受的工作簿基准单独比较；未来工作 |

---

## 2. 每一步的 agent 与 tools

工具名后的状态：`已实现` = `src/` 内有对应实现；`目标` = 尚未实现，定义在 `docs/issues/backlog.md`。

| Step | Agent | 职责（capability） | 主要 Tools | 状态 |
|---|---|---|---|---|
| 0 | `intake-agent` | 识别文件真伪与可解析性，阻断则报 error | `scan_ooxml_package` `OpenpyxlWorkbookReader.read_manifest`（`detect_format`/`decrypt_probe` 目标） | 部分 |
| **1** | **`decomposition-agent`** | **按内容类型全覆盖分类提取，不做语义判断** | 见下方 Step 1 工具图 | 部分 |
| **1b** | `completeness-agent` | 独立验证输出是否覆盖输入，产出 handoff | `verify_completeness`、`build_handoff`、`render_handoff_markdown` | 已实现 |
| 2 | `index-agent` | 为 Step 1 handoff 建索引并生成导航摘要 | `src/excel_to_act/steps/step2/workflow.py:build_index` 及其 `_write_index` 已实现，支持 `step1.v1` / `step1.batch.v1`，写入批次 `index.json` 和 `INDEX.md`；独立的 `summarize_artifacts` helper 仍为目标。旧版 `src/excel_to_act/store/local_store.py:LocalArtifactStore._write_index` 写入每工作簿的 `artifact_index.json`。 | 部分（独立摘要 helper 仍为目标） |
| 3 | `analysis-agent` | 模块分类、域标签、边界确认 | `classify/rules.py`+`classifier.py`、`confirm/templates.py`（已实现） | 部分 |
| 4 | `generation-agent` | 生成结构化 Python（保留溯源） | `emit_module` / `emit_package`（目标） | 目标 |
| 5 | `validation-agent` | 生成前比较已保存的工作簿缓存与实际重算结果（#7）；Step 4 后另行检查生成代码等价性 | 旧版 `inspect` 已有 `ValidationReport` 和 `formulas` 适配器；完整 Agent 流程与代码等价性检查仍是目标 | 部分实现 |

### Step 1 内部流程

```mermaid
flowchart LR
    subgraph STEP1 ["Step 1 · Lossless Decomposition"]
        direction TB
        PKG["scan_package<br/>OOXML 全量部件"]
        CEL["extract_cells<br/>值 + 公式 + 缓存值"]
        FRM["extract_formulas<br/>tokenizer → 引用"]
        NAM["extract_names<br/>defined names / LAMBDA"]
        STY["extract_styles<br/>numFmt / fill / font"]
        TBL["extract_tables<br/>ListObjects / pivot"]
        VLD["extract_validation<br/>下拉 / 模拟运算表"]
        CTL["extract_controls<br/>表单控件 linkedCell"]
        CMT["extract_comments"]
        EXT["extract_external<br/>外部链接 / 连接"]
        VBA["extract_vba<br/>oletools → 模块源码"]
        XLINK["extract_vba_cell_links<br/>VBA ↔ 单元格 依赖边"]
        OPA["extract_opaque<br/>图表 / 透视缓存 / PowerPivot"]
        COV["verify_coverage<br/>覆盖不变量校验"]
    end

    PKG --> CEL --> FRM --> COV
    PKG --> NAM --> COV
    PKG --> STY --> COV
    PKG --> TBL --> COV
    PKG --> VLD --> COV
    PKG --> CTL --> COV
    PKG --> CMT --> COV
    PKG --> EXT --> COV
    PKG --> VBA --> XLINK --> COV
    PKG --> OPA --> COV
```

旧 `inspect` 的 `CoverageSummary.discovered_workbook_objects` 来自独立 OOXML 源扫描，recognized 数量按匹配到的源身份统计。仅有汇总计数不能证明逐个身份都已覆盖；`verify_completeness` 会比较源与 inventory 身份、校验存储计数、保存差异诊断，并在漏项时以非零码退出（[#5 / PR #38](https://github.com/ferryhe/excel_to_act/pull/38)）。人工 `step1 coverage` 使用独立的逻辑对象和包部件账本。两条路径都不表示所有 Excel 功能均已语义解析。

---

## 3. Step 1 输出目录

### 3.1 当前实现

旧单工作簿 `inspect` CLI 是 `excel-to-act inspect <workbook> --out <dir>`。若输入路径是目录，`inspect` 会返回受控的读取错误并写出仅含诊断的失败运行产物；目录批次由上文新增的 `excel-to-act step1 convert` 处理。成功的 `inspect` 运行写出七个数据 JSON、`run_metadata.json` 和供人阅读的 `handoff.md`：

```text
<out>/workbooks/<workbook_sha256>/<run_id>/
  workbook_manifest.json
  inventory.json
  dependency_graph.json
  module_classification.json
  confirmation_template.json
  completeness.json
  handoff.json
  handoff.md
  run_metadata.json
<out>/workbooks/<workbook_sha256>/artifact_index.json
<out>/artifact_index.json
<out>/<artifact>               # 最新运行产物的便捷别名（包括 handoff.md）
```

仅诊断失败的运行会写出 `workbook_manifest.json`、`completeness.json`、`handoff.json`、`handoff.md` 和 `run_metadata.json`，以及索引；不会生成 `inventory.json`、`dependency_graph.json`、`module_classification.json` 或 `confirmation_template.json`。根目录别名只反映最新一次运行：失败运行会移除这四个缺失产物的旧别名，之前成功运行的产物仍保留在各自的规范运行目录中。

### 3.2 目标布局（Step 1）

- `--input` 支持**目录**（批量）
- 命名空间沿用 `<workbook_sha256>/<run_id>/`，**不用 slug**：slug 由文件名派生，同名不同内容会串、改名会断链
- 每一类同时产出 `.md`（人读）与 `.json`（机器读）
- 旧 `inspect` 流程的 `Handoff` 契约已实现并写出 JSON/Markdown；其 `CoverageSummary` 仍嵌在 `WorkbookInventory` 中，独立 coverage 产物属于旧目标布局。当前目录式 `step1` 使用本节所述独立检查与批次 handoff。

```text
output/step1_decomposition/workbooks/<workbook_sha256>/<run_id>/
  00_manifest.{md,json}         文件身份、sha256、格式、加密状态
  01_package.{md,json}          OOXML 全量部件 + content type + relation
  02_sheets/                    sheet 级：名称/顺序/可见性/类型/维度
  03_cells/                     单元格：值、公式、缓存值、number_format
  04_formulas/                  公式引用图（含 VBA 边）
  05_names/                     defined names、LAMBDA、隐藏名称
  06_tables/                    ListObjects、结构化引用、pivot 定义与缓存
  07_styles/                    numFmt / font / fill / border / dxfs
  08_conditional_formatting/
  09_validation/                数据验证 + 模拟运算表（dataTable）
  10_comments/                  批注与作者
  11_charts/
  12_external/                  externalLinks、connections、Power Query
  13_controls/                  表单控件 linkedCell、切片器、时间线
  14_vba/                       模块源码 + 过程清单 + 静态调用关系
  15_xlm_macros/                Excel 4.0 宏与 DDE
  99_opaque/                    图表二进制、Power Pivot、OLE、签名等
  coverage.{md,json}            旧 inspect 目标布局；当前 step1 coverage 另有独立逻辑对象/部件账本
  handoff.{md,json}             旧 inspect 目标布局；当前 step1 已写出双格式 handoff
```

`handoff.json` 是唯一被下游消费的契约：列出每类产物的路径、条数、schema 版本、以及 `opaque` 摘要。

---

## 4. Step 1 收尾：完备性检查与 handoff

分解结束后 **必须先做完备性检查，再产出 handoff**。这是 Step 1 的出口，也是下游 agent 唯一需要读的入口。

`verify/completeness.py` 的 `verify_completeness()` 不会只依赖 `WorkbookInventory.coverage` 汇总计数。`discovered_workbook_objects` 来自独立源扫描，但汇总计数不能证明逐个身份都匹配。该检查会从包中独立发现源对象、比较源与 inventory 身份、校验存储计数，并在对象缺失时失败：

| 检查 | 判据 | 严重级 |
|---|---|---|
| `sheets_accounted` | `workbook.xml` 里每个 `<sheet>` 都有对应 `SheetInventory` | error |
| `cells_accounted` | 每表单元格数与 worksheet XML 的独立扫描一致 | error |
| `content_parts_accounted` | 每个内容部件**有产出证据**（如 `xl/tables/` 必须存在 `table` 类型的 range）或已标 opaque | error |
| `formulas_linked` | 每个公式格有出边，或有 `formula_reference_parse` 记录 | warning |
| `metadata_parts_accounted` | `docProps/` 等尚未建模的元数据部件 | info |
| `coverage_arithmetic` | recognized/discovered 对象总数是否与存储的覆盖计数一致。各类型独立的 `*_identities` 检查负责匹配身份并发现缺失、多余或重复对象；`verify_completeness` 会执行两类检查。包部件单独核算。 | error |

> 部件是否"已采集"不看路径前缀，而看**产出里有没有对应证据**——否则哪天采集逻辑被删掉，检查也不会报警。

状态与后果：

- `fail`：任一 error 级未过 → **CLI 退出码 1**；产物照写，但必须视为不可用
- `warn`：仅 warning 级未过
- `pass`：全部通过

产物（run 目录 `<out>/workbooks/<sha256>/<run_id>/`；`handoff.md` 与 `handoff.json` 同时复制一份到 `<out>/` 根，便于直接找到）：

| 文件 | 给谁看 | 内容 |
|---|---|---|
| `completeness.json` | 机器 | 检查项、期望/实际、gaps、blocking_reasons |
| `handoff.json` | 机器 | `summary` 计数、产物清单（路径 / sha256 / 条数）、coverage、opaque 汇总、blockers、warnings、next_actions |
| `handoff.md` | **人** | 同一份信息的简短英文摘要，包含 `At a glance` / `Blockers` / `Warnings` / `Artifacts` / `Unresolved (opaque)` / `Next steps`。**不打开任何 JSON 就能判断本次分解能否被信任** |

`handoff.md` 顶部是工作簿速写（工作表 / 单元格 / 公式格 / 缓存值 / 已定义名称 / 表 / 合并区 / 模拟运算表 / 表单控件 / VBA 模块 / 依赖图节点与边），供人快速核对"是不是我那个文件、规模对不对"。

## 5. Step 1 全覆盖分类清单

**图例（三态）**

| 标记 | 含义 |
|---|---|
| `已采集` | 已进 `inventory.json` |
| `opaque` | 明确记录为未解析或不支持；Step 1 还会逐字节保存包部件 |
| **静默丢失** | 既未采集也未明确标记为 opaque —— 违反硬约束 |

### A. 包与文件层

| 内容 | OOXML 源 | 状态 |
|---|---|---|
| 文件身份（sha256/大小/格式） | — | 已采集 |
| 全量部件 + content type | `[Content_Types].xml` | 已采集 |
| 文档属性 | `docProps/*.xml` | **缺失**：应采集，不是标 opaque（可解析） |
| 自定义 XML / Power Query `DataMashup` | `customXml/` | opaque（token 已补，见 I 组）；M 代码本身未解析 |
| 数字签名 | `_xmlsignatures/` | opaque（token 已补，见 I 组） |
| 加密 / 损坏文件 | `EncryptedPackage` | **已作为受控失败处理**：PR #37 保存仅含诊断的失败运行产物，并让 CLI 以受控非零码退出。不执行解密（[Issue #8](https://github.com/ferryhe/excel_to_act/issues/8)；[PR #37](https://github.com/ferryhe/excel_to_act/pull/37)） |

### B. 工作簿结构层

| 内容 | 状态 |
|---|---|
| sheet 名称/顺序/可见性/维度 | 已采集 |
| `calcMode`（manual/auto） | 已采集（`WorkbookManifest.calc_mode`） |
| `calcPr` 其余（iterate / iterateCount / iterateDelta / calcId / fullCalcOnLoad） | **静默丢失** —— 循环迭代设置，直接关系循环引用怎么解 |
| `xl/calcChain.xml`（最后计算顺序） | Step 1 会记录为 opaque 并逐字节保存；不会解析计算顺序 |
| defined names | 工作簿级和工作表级名称都会采集；区域目标会展开为不带工作表前缀的 A1 地址（每个区域一行），inventory 元数据保留声明作用域。Step 1 源清单按声明计数并保留原始名称文本。`hidden` 标志未采集 |
| LAMBDA / 名称里定义的自定义函数 | Step 1 会保留原始名称文本，但**尚未判定 LAMBDA**，也未解析名称公式里的引用 |

### C. 单元格内容层

| 内容 | 状态 |
|---|---|
| 值 / 公式 / 错误值（`#N/A`、`#REF!`） | 已采集（错误值以字面量存） |
| number_format、data_type、style_id | 已采集 |
| 空单元格 | **未采集**：`extractor.py:55` 直接 `continue`；`CellKind.blank` 是死枚举（要么采集，要么删枚举） |
| 缓存值（`data_only=True`） | **已采集并验收**（[PR #15](https://github.com/ferryhe/excel_to_act/pull/15) 与 [PR #16](https://github.com/ferryhe/excel_to_act/pull/16) 已合入；[Issue #4](https://github.com/ferryhe/excel_to_act/issues/4) 已关闭）：`cached_value` + `cached_value_available`；整簿无缓存值时产出 `missing_cached_values` warning。对账 oracle 的基石 |
| `cell.formula_attributes` | Step 1 会在 `source_facts.json` 和 `CellInventory` 中保存原始公式文本与完整 `<f>` 属性，包括 shared、array 和 `dataTable` 属性。旧 `inspect` 不填充这些原始公式事实；可空的存在性和日期字段仍为未知。保存属性不等于解释动态数组 spill 行为 |
| 富文本 runs、phonetic | **静默丢失** |

### D. 样式与呈现层（识别 input/output 的关键信号）

| 内容 | 状态 |
|---|---|
| 合并单元格 | 已采集 |
| 行高列宽、隐藏行列、冻结窗格 | 已采集 |
| 保护（sheet 级） | 已采集 |
| 批注（文本+作者）、超链接（target+location） | 已采集 |
| 条件格式 | 已采集**范围**；`cfRule` 类型/dataBar/colorScale/iconSet 未采集 |
| fonts / fills / borders / numFmts / dxfs | **静默丢失** |
| 打印设置、页眉页脚 | **静默丢失** |

### E. 结构化对象层

| 内容 | 状态 |
|---|---|
| 数据验证 | 已采集：范围 + `type`/`formula1`/`formula2`；缺 `operator`/`allowBlank`/`showDropDown`/prompt+error 文案，且 list 型来源未解析成区域 |
| Excel Tables（ListObjects） | 已采集 `ref`+`name`；**列定义、totalsRow、排序/筛选态未采集** |
| 透视表定义 + 缓存记录 | opaque（`/pivot` 在 `OPAQUE_MARKERS` 内） |
| `xl/connections.xml` | opaque（`/connections` 在内） |
| 切片器 / 时间线（`xl/slicers/`、`xl/timelines/`） | Step 1 会记录为 opaque 并逐字节保存；尚未解析其语义内容 |
| Power Query `DataMashup` | Step 1 会将 `customXml/` 部件记录为 opaque 并逐字节保存；M 代码未解析 |

### F. 图形与控件层

| 内容 | 状态 |
|---|---|
| 图表（含 `SERIES()` 引用） | opaque（`/charts/` 在内） |
| 图片 / 形状 / 文本框的单元格链接 | opaque（`/drawings/`、`/media/` 在内） |
| OLE 嵌入对象 | opaque（`/embeddings/` 在内） |
| **表单控件 `linkedCell`**（`vmlDrawing` + `ctrlProps`） | Step 1 会从 `xl/drawings/vmlDrawing*.vml` 提取有绑定的 `ClientData`。只有每个 shape 都有一个受支持且已绑定的 `ClientData` 时，该 VML 部件才算非 opaque；Note/Pict、未绑定或缺少 `ClientData`，以及其它 VML 图形对象会使整个部件标为 opaque。混合部件中的受支持绑定仍会提取，原部件会逐字节保存 |
| **模拟运算表 `dataTable`** | Step 1 会从 worksheet XML 直接读取 `<f t="dataTable" ref r1 r2>`，并生成 `data_table` 范围对象和输入格事实。它保存源属性，但不会计算公式或重新计算模拟运算结果 |

### G. 代码与自动化层

`工具` 列区分 `依赖库` 与 `自研解析器`。

| 内容 | 源 | 工具 | 状态 |
|---|---|---|---|
| **VBA 模块源码** | `xl/vbaProject.bin` | 依赖库 `oletools`(olevba) `>=0.60.2`，BSD-3（核实于 2026-10-02，PyPI 最新 0.60.2 / 2024-07-02） | **当前分支已实现提取路径**（本地任务 `LOCAL-VBA`，见 [PR #15](https://github.com/ferryhe/excel_to_act/pull/15)）：`extract_vba_project` 产出模块源码 + 过程名；未装 oletools 则降级 warning，不崩 |
| **VBA ↔ 单元格 依赖边** | 上一步源码 | 自研 `extract_vba_cell_links` → `build_vba_edges` | **已采集**：`FormulaGraph` 追加 `relationship="vba_ref"` 边，与公式图同一节点命名空间（新增 `GraphNodeKind.vba` 与 `GraphEdge.confidence`） |
| **Excel 4.0 宏（XLM）** | `xl/macrosheets/` | 依赖库 oletools `olevba`（XLM 走另一入口，需可选依赖 `XLMMacroDeobfuscator`） | Step 1 会记录对应包部件为 opaque 并逐字节保存；不会提取宏内容 |
| **DDE 链接** | — | 依赖库 oletools `msodde` | **静默丢失** |
| **LAMBDA / 自定义名称函数** | defined names | 自研 `extract_names` | 名称已采集，判定缺失 |
| **RTD / 加载项** | `volatileDependencies.xml`、`webExtensions/` | 自研 `scan_package` | Step 1 会记录对应包部件为 opaque 并逐字节保存；不会解释加载项行为 |

**为什么 VBA 必须提取源码**：精算老模型普遍用 `Range("B7")` / `Names("Mort_qx")` 驱动计算，这类依赖边在公式图里完全不可见，只扫公式会得到一张**断裂的依赖图**（`graph/builder.py` 目前完全不看 VBA）。

落地约束（写死，避免返工）：

1. 抽出的边**必须写进 `FormulaGraph`**，用 `GraphEdge.relationship = "vba_ref"`，**禁止另造契约或新边类型**——否则 `dependency_graph.json` 出现两套节点空间，引用关系无法闭合
2. 字面量匹配会漏：`Cells(r, c)`、`"B" & i` 拼接、经变量/命名区域间接寻址。抽取结果必须带 `confidence`，无法静态确定的记为 `vba_ref_unresolved` 并进 confirmation，不得假装确定性

### H. 语义信号层

| 内容 | 状态 |
|---|---|
| 批注 | 已采集 |
| 数据验证下拉来源 | 已采集 `formula1`，未解析成区域 |
| 模拟运算表 | Step 1 按 F 组所述采集；不会重新计算结果 |
| 分组/大纲、Excel 场景管理器 | **静默丢失** |
| Solver / Goal Seek（存于 `Solver_*` 隐藏名称 + VBA） | **静默丢失** |

### I. 被记录为 opaque 的包部件

`OPAQUE_MARKERS` 现在包含以下部件。Step 1 会将它们记录为 opaque 并保存其字节；这些标记不表示已经做了语义解析。它们不是当前的静默丢失项：

```text
xl/model/                     Power Pivot 数据模型（ABF 二进制）
_xmlsignatures/               数字签名
xl/calcChain.xml              计算链
xl/ctrlProps/                 表单控件属性（linkedCell）
xl/slicers/  xl/timelines/    切片器 / 时间线
xl/macrosheets/               Excel 4.0 宏
customXml/                    Power Query / 自定义 XML 映射
xl/volatileDependencies.xml   RTD
webExtensions/                Office 加载项
```

文档属性应直接采集，而不应标为 opaque：

```text
docProps/*.xml                文档属性（可解析）
```

模拟运算表已如 F 组所述从 worksheet XML 采集。标记为 opaque 只保证保存，不保证可用。后续可继续补充 A–H 中剩余的语义解析。Step 1 已读取公式缓存值和受支持的 VML 控件绑定。

### J. 不可解析（记为 opaque，不尝试解释）

Power Pivot 数据模型、透视缓存二进制、图表渲染、嵌入媒体、加密内容、签名。

### K. 规划词汇与已实现的运行时契约

Phase 1 计划中的模块类别和精算提示列表是提案/规划词汇，并不要求运行时枚举包含这些成员。当前已实现的 `ModuleCategory` 契约有 9 个值（`input`、`data_table`、`formula_block`、`lookup_block`、`output`、`presentation`、`external_dependency`、`unsupported_opaque`、`other`）；已实现的 `ActuarialHint` 契约有 6 个值（`assumption`、`rate_table`、`cashflow`、`projection`、`output`、`unknown`）。定义见 `src/excel_to_act/schemas/artifacts.py:273-292`。计划在 `docs/plans/phase1_excel_decomposition_plan.md:67-80` 列出 12 个拟议模块标签，并在 `docs/plans/phase1_excel_decomposition_plan.md:88-97` 列出 10 个拟议提示标签；这些差异不表示运行时枚举缺少成员。

---

## 6. 当前代码布局与目录路线图

当前 Step 1 的工作流、只用于 Step 1 的 OOXML 源扫描器和随安装包分发的 agent 定义放在同一个业务步骤包中：

```text
src/excel_to_act/
  steps/
    __init__.py
    step1/
      __init__.py
      workflow.py
      source_scan.py
      agent.md
  interfaces/                 CLI 入口，复用 Step 1 工作流
  ingest/                     共享文件读取与 OOXML 部件能力
  inventory/ verify/ graph/   共享 inventory 与检查/图能力
  classify/ confirm/          共享分类与确认能力
  schemas/ store/ report/     共享数据契约、存储与报告能力
  plugins/ orchestrator/      共享工具契约/旧 Phase 1 编排
tests/
```

每个已实现的人工业务步骤在 `steps/<step>/` 中拥有自己的 agent 定义，并可按职责拆成多个实现文件；Step 1 将转换工作流与源扫描器分在两个 Python 文件中。步骤包直接复用现有共享模块，不复制 shared capability，也不为尚未实现的步骤创建空包。

### 先前提出的目标布局（仅供路线图参考）

下方目录草图及说明保留先前的路线图内容，描述的是未来提案，不是当前源代码树。

```text
excel_to_act/
  src/excel_to_act/          库核心（分层不变；规范性分层定义见 docs/design/layered_architecture.md）
    interfaces/              CLI + API + agent 入口
    orchestrator/            step 状态机
    plugins/                 可替换工具契约与注册表
    ingest/ inventory/ graph/ classify/ confirm/ store/ report/ schemas/
    tools/                   —— Step 1 工具登记目录（可导入、随包分发）
      step0_intake/  step1_decomposition/  step2_index/ ...
    validation/              —— 目标流程：旧版 #7 比较位于 verify/numerical.py；代码等价性留待后续
  agents/                    每步 agent 定义：能力 / 可用 tools / IO 契约
    step1_decomposition.agent.md ...
  skills/                    可安装给外部 agent 的 skill 包
    excel-to-act/SKILL.md
  input/                     默认输入目录（CLI 可指向任意目录；已 gitignore）
  output/                    按 step 分目录（已 gitignore）
    step1_decomposition/ step2_index/ step3_analysis/ step4_generation/ step5_validation/
  schemas/ docs/ examples/ tests/
```

要点：

- 可打包的能力代码**放在 `src/excel_to_act/` 下**，而不是仓库根级目录：`pyproject.toml` 的 `packages.find where = ["src"]` 不会打包根级 Python 包。`plugins/contracts.py` 的现有 6 个 Protocol 保持不变；#7 的单一 `formulas` 适配器由 `verify/numerical.py` 直接调用。
- `plugins/registry.py` 的全局 `registry` 目前**仅测试使用**，生产模块尚未注册 —— "注册进 registry"仍是目标态
- `agents/` 里每个 agent 文件写死：能力边界、可调用 tools、输入/输出 artifact、禁止事项
- `skills/` 需要 `package-data`/`include-package-data` 才会随 wheel 分发，当前无此配置，待补
- `input/`、`output/` 只是**默认落点**，运行时由 CLI 参数决定

---

## 7. Phase 边界

- 人工 Step 1/2 执行源文件分解、索引、确定性视图和证据导航，不做语义判断。旧版 `inspect` 则单独执行启发式分类并生成确认问题。Phase 1.5 由 [Issue #3](https://github.com/ferryhe/excel_to_act/issues/3) 跟踪；本 PR 对齐 backlog 并在合入后关闭该 Epic。
- README Step 3 是更广泛的语义分析目标。当前旧版分类只使用[分层架构](docs/design/layered_architecture.md)中定义的运行时枚举，不代表已确认的领域含义。
- 旧版 `inspect` 已将已保存的公式缓存与独立的可选 `formulas` 重算结果比较，并记录缓存新鲜度未知及未覆盖单元格。Step 4 之后的生成代码等价性仍是后续工作。
- 旧版 `inspect` 按工作簿和 run ID 写出 Phase 1 产物（见 §3.1）；人工 Step 1/2 使用独立的源文件与 reading package 契约。两者均未实现 §3.2 的目标目录树。

## 参见

- `docs/plans/phase1_excel_decomposition_plan.md`（Phase 1 方案，≈ 本文件 Step 0–3）
- `docs/plans/pr_plan_phase1.md`（PR 切分与验收）
- `docs/research/excel_tooling_survey.md`（A1 · 现成库选型结论）
- `docs/issues/backlog.md`（Phase 1.5 Epic 收尾与 Issue 源）
