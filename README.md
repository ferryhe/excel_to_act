# excel_to_act

Excel-to-actuarial-model tooling：**CLI + API + agent-ready**，可按 skill 安装给任意 agent 使用。

第一原则：**先做无损分解（lossless decomposition），再做理解与生成**。任何读不懂的东西必须记为 opaque，禁止静默丢弃。

```text
recognized_inventory_objects + unsupported_or_opaque_objects = discovered_workbook_objects
```

> ⚠️ 注意：上述不变量**当前尚未真正生效**——`inventory/extractor.py:100` 把 `discovered` 直接写成 `recognized + opaque`，是恒等式；且一批部件既未采集也未标 opaque（见 §4 的"静默丢失"）。修复见 `docs/issues/backlog.md` issue #5 与 §4 I。

---

## 1. 流水线总览

```mermaid
flowchart TD
    IN["input/ · 目录<br/>目标 CLI: --input<br/>当前 CLI: inspect &lt;workbook&gt; --out"]
    S0["Step 0 · Intake &amp; Identify<br/>识别格式 / 加密 / 包完整性"]
    S1["Step 1 · Lossless Decomposition<br/>按内容类型全覆盖分类提取"]
    O1[("output/step1_decomposition/<br/>按种类 md + json")]
    V["verify_completeness<br/>独立重算对象全集并比对"]
    H["handoff.json + handoff.md<br/>Step 1 唯一出口"]
    S2["Step 2 · Artifact Index<br/>为后续 agent 建产物索引与摘要"]
    O2[("output/step2_index/<br/>INDEX.md + index.json")]
    S3["Step 3 · Semantic Analysis<br/>模块分类 / 域标签 / 结构规约"]
    O3[("output/step3_analysis/")]
    S4["Step 4 · Structured Python Generation"]
    O4[("output/step4_generation/")]
    S5["Step 5 · Validation &amp; Reconciliation<br/>多 oracle 数值对账"]
    O5[("output/step5_validation/")]

    IN --> S0 --> S1 --> O1 --> V --> H --> S2 --> O2 --> S3 --> O3 --> S4 --> O4 --> S5 --> O5

    style S1 fill:#dff5e1,stroke:#2e7d32,stroke-width:2px
    style O1 fill:#eef7ff,stroke:#1565c0
```

**落地状态**：Step 0–2 已部分落地（`ingest/` `inventory/` `graph/` `classify/` `confirm/` `store/` `orchestrator/` `interfaces/`；Step 2 的 `artifact_index.json` 已由 `store/local_store.py` 产出）。Step 3 属 plan 的 Phase 1 范围、已部分落地（`classify/`）。Step 1 的 JSON 与 Markdown handoff 已实现；**Step 4–5 未开工**。

**当前建造目标：Step 1 的无损分解扩展**（见 §4 的缺口清单）。

### 与 `phase1_excel_decomposition_plan.md` 的映射

| README 流水线 | plan 中的位置 |
|---|---|
| Step 0–3 | **plan Phase 1 范围内**（ingest → inventory → graph → classify → confirm → store → report） |
| Step 4（生成 Python） | plan 「Non-goals」明确排除 |
| Step 5（对账） | plan 仅提到 LibreOffice/xlwings 未来 oracle，未排期 |

---

## 2. 每一步的 agent 与 tools

工具名后的状态：`已实现` = `src/` 内有对应实现；`目标` = 尚未实现，定义在 `docs/issues/backlog.md`。

| Step | Agent | 职责（capability） | 主要 Tools | 状态 |
|---|---|---|---|---|
| 0 | `intake-agent` | 识别文件真伪与可解析性，阻断则报 error | `scan_ooxml_package` `OpenpyxlWorkbookReader.read_manifest`（`detect_format`/`decrypt_probe` 目标） | 部分 |
| **1** | **`decomposition-agent`** | **按内容类型全覆盖分类提取，不做语义判断** | 见下方 Step 1 工具图 | 部分 |
| **1b** | `completeness-agent` | 独立验证输出是否覆盖输入，产出 handoff | `verify_completeness`、`build_handoff`、`render_handoff_markdown` | 已实现 |
| 2 | `index-agent` | 为 Step 1 产物建索引与摘要 | `_write_index`（已实现）；`build_index`/`summarize_artifacts` 目标 | 部分 |
| 3 | `analysis-agent` | 模块分类、域标签、边界确认 | `classify/rules.py`+`classifier.py`、`confirm/templates.py`（已实现） | 部分 |
| 4 | `generation-agent` | 生成结构化 Python（保留溯源） | `emit_module` / `emit_package`（目标） | 目标 |
| 5 | `validation-agent` | 多 oracle 数值回归与对账 | `cached_value` / `formulas_oracle` / `libreoffice_oracle`（目标，issue #7） | 目标 |

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

`verify_coverage` **尚不存在**：`CoverageSummary.discovered` 目前是 `recognized + opaque` 的恒等式，只有等 issue #5 落地后 `coverage` 三数才有意义。

---

## 3. Step 1 输出目录

### 3.1 当前实现

CLI 是 `excel-to-act inspect <workbook> --out <dir>`，且 `dir_okay=False`——**当前不接受目录入参**。每次运行写出七个数据 JSON、`run_metadata.json` 和供人阅读的 `handoff.md`：

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

根目录会复制最新运行的产物作为便捷别名；`handoff.md` 和 `handoff.json` 也会复制到根目录。

### 3.2 目标布局（Step 1）

- `--input` 支持**目录**（批量）
- 命名空间沿用 `<workbook_sha256>/<run_id>/`，**不用 slug**：slug 由文件名派生，同名不同内容会串、改名会断链
- 每一类同时产出 `.md`（人读）与 `.json`（机器读）
- `Handoff` 契约已实现并写出 JSON/Markdown；`CoverageSummary` 仍嵌在 `WorkbookInventory` 中，独立 coverage 产物仍是目标布局的一部分

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
  coverage.{md,json}            recognized / opaque / discovered（待新契约）
  handoff.{md,json}             交接文件：下一步 agent 的入口（待新契约）
```

`handoff.json` 是唯一被下游消费的契约：列出每类产物的路径、条数、schema 版本、以及 `opaque` 摘要。

---

## 4. Step 1 收尾：完备性检查与 handoff

分解结束后 **必须先做完备性检查，再产出 handoff**。这是 Step 1 的出口，也是下游 agent 唯一需要读的入口。

`verify/completeness.py` 的 `verify_completeness()` **不复用** `WorkbookInventory.coverage`——它的 `discovered` 是 `recognized + opaque` 的恒等式，证明不了任何事。而是**直接从包里重算对象全集**（worksheet XML 数非空单元格 + 枚举包部件），再与产出比对：

| 检查 | 判据 | 严重级 |
|---|---|---|
| `sheets_accounted` | `workbook.xml` 里每个 `<sheet>` 都有对应 `SheetInventory` | error |
| `cells_accounted` | 每表单元格数与 worksheet XML 的独立扫描一致 | error |
| `content_parts_accounted` | 每个内容部件**有产出证据**（如 `xl/tables/` 必须存在 `table` 类型的 range）或已标 opaque | error |
| `formulas_linked` | 每个公式格有出边，或有 `formula_reference_parse` 记录 | warning |
| `metadata_parts_accounted` | `docProps/` 等尚未建模的元数据部件 | info |
| `coverage_arithmetic` | `recognized + opaque` 是否等于独立重算的 `discovered` | info |

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
| `handoff.md` | **人** | 同一份信息的简要中文摘要：一眼看懂 / ⛔ 阻塞 / ⚠️ 警告 / 产出 / 未解析 / 下一步。**不打开任何 JSON 就能判断本次分解能否被信任** |

`handoff.md` 顶部是工作簿速写（工作表 / 单元格 / 公式格 / 缓存值 / 已定义名称 / 表 / 合并区 / 模拟运算表 / 表单控件 / VBA 模块 / 依赖图节点与边），供人快速核对"是不是我那个文件、规模对不对"。

## 5. Step 1 全覆盖分类清单

**图例（三态）**

| 标记 | 含义 |
|---|---|
| `已采集` | 已进 `inventory.json` |
| `opaque` | 有记录、无解析（进 `unsupported_features`，符合不变量） |
| **静默丢失** | 既未采集也未标记 —— **违反硬约束，必须补**（见 I 组） |

### A. 包与文件层

| 内容 | OOXML 源 | 状态 |
|---|---|---|
| 文件身份（sha256/大小/格式） | — | 已采集 |
| 全量部件 + content type | `[Content_Types].xml` | 已采集 |
| 文档属性 | `docProps/*.xml` | **缺失**：应采集，不是标 opaque（可解析） |
| 自定义 XML / Power Query `DataMashup` | `customXml/` | opaque（token 已补，见 I 组）；M 代码本身未解析 |
| 数字签名 | `_xmlsignatures/` | opaque（token 已补，见 I 组） |
| 加密 / 损坏文件 | `EncryptedPackage` | **已确认违反硬约束**：`ingest/ooxml_package.py:51` 的 `ZipFile()` 无异常保护，会抛未捕获异常（issue #8，P1）。要求：报 error，不 panic |

### B. 工作簿结构层

| 内容 | 状态 |
|---|---|
| sheet 名称/顺序/可见性/维度 | 已采集 |
| `calcMode`（manual/auto） | 已采集（`WorkbookManifest.calc_mode`） |
| `calcPr` 其余（iterate / iterateCount / iterateDelta / calcId / fullCalcOnLoad） | **静默丢失** —— 循环迭代设置，直接关系循环引用怎么解 |
| `xl/calcChain.xml`（最后计算顺序） | **静默丢失** |
| defined names | 已采集，**仅工作簿级**；sheet 级名称（`ws.defined_names`）与 `hidden` 标志未采集 |
| LAMBDA / 名称里定义的自定义函数 | 已采集名称本体，**未做 LAMBDA 判定**，也未解析名称公式里的引用 |

### C. 单元格内容层

| 内容 | 状态 |
|---|---|
| 值 / 公式 / 错误值（`#N/A`、`#REF!`） | 已采集（错误值以字面量存） |
| number_format、data_type、style_id | 已采集 |
| 空单元格 | **未采集**：`extractor.py:55` 直接 `continue`；`CellKind.blank` 是死枚举（要么采集，要么删枚举） |
| 缓存值（`data_only=True`） | **已采集**（issue #4 已完成）：`cached_value` + `cached_value_available`；整簿无缓存值时产出 `missing_cached_values` warning。对账 oracle 的基石 |
| `cell.formula_attributes` | **静默丢失**：shared formula 的 `si`/`ref`、array formula 的 `t=array`/`ref`、动态数组 spill。注：公式**展开**由 openpyxl 完成，不是缺口；数组公式的**文本**现已保留（不再被 `str()` 成对象 repr），但 `ref` 与 spill 范围仍缺 |
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
| 切片器 / 时间线（`xl/slicers/`、`xl/timelines/`） | **静默丢失** |
| Power Query `DataMashup` | **静默丢失** |

### F. 图形与控件层

| 内容 | 状态 |
|---|---|
| 图表（含 `SERIES()` 引用） | opaque（`/charts/` 在内） |
| 图片 / 形状 / 文本框的单元格链接 | opaque（`/drawings/`、`/media/` 在内） |
| OLE 嵌入对象 | opaque（`/embeddings/` 在内） |
| **表单控件 `linkedCell`**（`vmlDrawing` + `ctrlProps`） | **已采集**（issue #16 已完成）：来自 `xl/drawings/vmlDrawing*.vml` 的 `ClientData`；`linked_cell` 解析为 `Sheet!A1` 并计入 `recognized`；注释型 `ClientData(Note)` 已排除 |
| **模拟运算表 `dataTable`** | **已采集**（issue #15 已完成）：存于单元格 `<f t="dataTable" ref r1 r2>`，openpyxl 会**丢掉公式文本**且 `str(cell.value)` 会损坏该格，故走 XML 直读；产出 `data_table` 范围对象 + 输入格 |

### G. 代码与自动化层

`工具` 列区分 `依赖库` 与 `自研解析器`。

| 内容 | 源 | 工具 | 状态 |
|---|---|---|---|
| **VBA 模块源码** | `xl/vbaProject.bin` | 依赖库 `oletools`(olevba) `>=0.60.2`，BSD-3（核实于 2026-10-02，PyPI 最新 0.60.2 / 2024-07-02） | **已采集**（issue #17 已完成）：`extract_vba_project` 产出模块源码 + 过程名；未装 oletools 则降级 warning，不崩 |
| **VBA ↔ 单元格 依赖边** | 上一步源码 | 自研 `extract_vba_cell_links` → `build_vba_edges` | **已采集**：`FormulaGraph` 追加 `relationship="vba_ref"` 边，与公式图同一节点命名空间（新增 `GraphNodeKind.vba` 与 `GraphEdge.confidence`） |
| **Excel 4.0 宏（XLM）** | `xl/macrosheets/` | 依赖库 oletools `olevba`（XLM 走另一入口，需可选依赖 `XLMMacroDeobfuscator`） | **静默丢失** |
| **DDE 链接** | — | 依赖库 oletools `msodde` | **静默丢失** |
| **LAMBDA / 自定义名称函数** | defined names | 自研 `extract_names` | 名称已采集，判定缺失 |
| **RTD / 加载项** | `volatileDependencies.xml`、`webExtensions/` | 自研 `scan_package` | **静默丢失** |

**为什么 VBA 必须提取源码**：精算老模型普遍用 `Range("B7")` / `Names("Mort_qx")` 驱动计算，这类依赖边在公式图里完全不可见，只扫公式会得到一张**断裂的依赖图**（`graph/builder.py` 目前完全不看 VBA）。

落地约束（写死，避免返工）：

1. 抽出的边**必须写进 `FormulaGraph`**，用 `GraphEdge.relationship = "vba_ref"`，**禁止另造契约或新边类型**——否则 `dependency_graph.json` 出现两套节点空间，Step 5 对账无法闭合
2. 字面量匹配会漏：`Cells(r, c)`、`"B" & i` 拼接、经变量/命名区域间接寻址。抽取结果必须带 `confidence`，无法静态确定的记为 `vba_ref_unresolved` 并进 confirmation，不得假装确定性

### H. 语义信号层

| 内容 | 状态 |
|---|---|
| 批注 | 已采集 |
| 数据验证下拉来源 | 已采集 `formula1`，未解析成区域 |
| 模拟运算表 | 见 F 组（**静默丢失**） |
| 分组/大纲、Excel 场景管理器 | **静默丢失** |
| Solver / Goal Seek（存于 `Solver_*` 隐藏名称 + VBA） | **静默丢失** |

### I. 当前**静默丢失**的部件总表（最高优先级，直接违反不变量）

`OPAQUE_MARKERS` 原有 8 个 token。以下部件不在其列，因此既不采集也不标 opaque，直接消失。**已补齐 token**（`ingest/ooxml_package.py`），现共 18 个：

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

另有两项**不应**标 opaque，而应直接采集：

```text
docProps/*.xml                文档属性（可解析）
<dataTable>                   模拟运算表（在 worksheet XML 内，openpyxl 丢弃）
```

**后续动作**：标 opaque 只保证"不丢"，不保证"可用"。按 A–H 各组把高价值项（缓存值、`dataTable`、`ctrlProps`、VBA）从 opaque 逐个升级为已采集。

### J. 不可解析（记为 opaque，不尝试解释）

Power Pivot 数据模型、透视缓存二进制、图表渲染、嵌入媒体、加密内容、签名。

### K. 已知契约漂移（不属于内容分类，但影响"全覆盖"的可信度）

| 定义来源 | 定义数 | 代码实现 | 差异 |
|---|---|---|---|
| `phase1_excel_decomposition_plan.md:67-80` module category | 12 | `schemas/artifacts.py:180-189` 9 个 | 缺 `workbook_meta` / `sheet_structure` / `calculation_chain` / `macro_or_code` |
| `phase1_excel_decomposition_plan.md:88-97` actuarial hint | 10 | `artifacts.py:192-199` 6 个 | 缺 `expense` / `claim_or_benefit` / `reserve` / `discount_curve` |

已记入 `docs/issues/backlog.md`，本清单不重复展开。

---

## 6. 目录安排

```text
excel_to_act/
  src/excel_to_act/          库核心（分层不变；L0-L3 定义待 issue #11）
    interfaces/              CLI + API + agent 入口
    orchestrator/            step 状态机
    plugins/                 可替换工具契约与注册表
    ingest/ inventory/ graph/ classify/ confirm/ store/ report/ schemas/
    tools/                   —— Step 1 工具登记目录（可导入、随包分发）
      step0_intake/  step1_decomposition/  step2_index/ ...
    validation/              —— Step 5 oracle runners（待 issue #7）
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

- `tools/` **放在 `src/excel_to_act/` 下**，而不是仓库根级目录：`pyproject.toml` 的 `packages.find where = ["src"]` 不会打包根级 Python 包。每个 tool 实现 `plugins/contracts.py` 的现有 6 个 Protocol（`WorkbookReader`/`InventoryExtractor`/`GraphBuilder`/`Classifier`/`ConfirmationBuilder`/`ArtifactStore`）；`OracleRunner` 待 issue #7 新增
- `plugins/registry.py` 的全局 `registry` 目前**仅测试使用**，生产模块尚未注册 —— "注册进 registry"仍是目标态
- `agents/` 里每个 agent 文件写死：能力边界、可调用 tools、输入/输出 artifact、禁止事项
- `skills/` 需要 `package-data`/`include-package-data` 才会随 wheel 分发，当前无此配置，待补
- `input/`、`output/` 只是**默认落点**，运行时由 CLI 参数决定

---

## 7. Phase 边界

- Step 0–2：只做分解、索引、摘要，**不做语义决策**
- Step 3（≈ plan Phase 1 的 classify）允许判断"这是什么模块"，且必须可确认、可覆盖（经 `confirm/`）
- 生成 Python（Step 4）之前必须通过对账（Step 5）
- 当前 Phase 1 按工作簿与运行 ID 写出七个数据 JSON、`run_metadata.json`、`handoff.md`，并提供根目录别名（见 §3.1）；**不是** §3.2 的分目录布局

## 参见

- `docs/plans/phase1_excel_decomposition_plan.md`（Phase 1 方案，≈ 本文件 Step 0–3）
- `docs/plans/pr_plan_phase1.md`（PR 切分与验收）
- `docs/research/excel_tooling_survey.md`（A1 · 现成库选型结论）
- `docs/issues/backlog.md`（issue 源，三段式：做什么 / 提交物 / 怎么检查）
