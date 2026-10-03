# A1 · Excel 工具调研：计算引擎与"Excel → 代码"类项目

- **Status:** Draft v2（已合并独立审核意见，待最终确认）
- **数据核实日期:** 2026-10-02
- **数据来源:** 各项目 PyPI 元数据页、GitHub README / 源码
- **文档类型:** 调研（Research）
- **审核:** 已由独立 subagent 对照 `docs/plans/*`、`pyproject.toml`、`src/**` 全量复核

## 关联文档（路径 + 当前状态）

| 编号 | 路径 | 状态 |
|---|---|---|
| A2 | `docs/research/docling_fidelity_assessment.md` | 未创建 |
| A3 | `docs/research/llm_table_encoding.md` | 未创建 |
| B1 | `docs/design/layered_architecture.md` | 未创建 |
| B2 | `docs/design/agent_reading_contract.md` | 未创建 |
| B3 | `docs/design/fidelity_rules.md` | 未创建 |
| C1 | `docs/design/views_l1_compiler.md` | 未创建 |
| E1/E2 | `docs/experiments/*` | 未创建 |
| F1/F2 | `docs/adr/*` | 未创建 |

---

## 0. 现状对齐（As-is）

本文是**已开工项目**的选型复核，不是立项前调研。仓库现有能力：

| 模块 | 文件 | 状态 | 本文结论相关性 |
|---|---|---|---|
| 读取 manifest | `src/excel_to_act/ingest/openpyxl_reader.py` | 已实现 | §4.6 确认 |
| OOXML 包扫描 | `src/excel_to_act/ingest/ooxml_package.py` | 已实现 | §4.6b（v1 漏评，已补） |
| 单元格清单 | `src/excel_to_act/inventory/extractor.py` | 已实现 | §4.6 确认；**缺 cached_value**，见 §7 |
| 依赖图 | `src/excel_to_act/graph/builder.py` | 已实现（正则解析） | §6.5 判定需替换 |
| 分类/确认/存储/编排/CLI | `classify/`, `confirm/`, `store/local_store.py`, `orchestrator/phase1.py`, `interfaces/cli.py` | 已实现 | 不在本文范围（归 B1） |
| 报告 | `src/excel_to_act/report/` | **仅空壳 `__init__.py`** | 不在本文范围（归 B1 / D2） |
| 插件协议 | `src/excel_to_act/plugins/contracts.py` | 6 个 Protocol，**无 OracleRunner** | §7 L3 需新增 |
| 可选依赖 | `pyproject.toml [project.optional-dependencies].formula` | `formulas>=1.3` + `xlcalculator>=0.5` 捆绑 | §7 建议拆分 |

---

## 1. 边界（Scope）

### In scope

评估 **"能读懂 Excel 公式并把它们算出来 / 编译成代码"** 的一类工具（计算引擎 / 编译器 / 求值器）+ **输入形态**相关工具，回答：

> 有没有现成的开源库，可以直接作为 excel_to_act 的 **核心** 依赖？如果没有，各自能放在哪一层？

### Out of scope

| 不在本文范围 | 归属 |
|---|---|
| Docling / markitdown / pandas 渲染等"文档解析器"是否保真 | **A2** |
| 给 LLM 的编码/压缩/采样策略 | **A3** |
| 分层架构的**正式定义**与层间契约 | **B1**（本文 §7 仅为输入 B1 的**非规范性初稿**） |
| 是否采用的最终决策记录 | **F1** |
| **新增**工具的集成代码与 PR 拆分 | C 类模块设计 / `docs/plans/pr_plan_phase1.md`（已落地实现以 `src/` 为准） |
| 性能基准 | E3（待建） |
| 测试策略与 fixture | `tests/README.md` + D 类实施计划 |
| 许可证合规的 CI 落地 | F1 ADR + CI workflow |
| Python 版本与运行环境矩阵 | `pyproject.toml` + `.github/workflows/ci.yml` |

---

## 2. 评估维度

| 维度 | 为什么重要 |
|---|---|
| **核心能力** | 是"求值"、"产出可读代码"还是"纯读取" |
| **函数覆盖** | 精算模型常用 LOOKUP / INDEX / MATCH / OFFSET / 财务函数的可得性 |
| **产物形态** | 是否产出可审计、可版本化的中间表示 |
| **结构覆盖** | defined name / 结构化引用 / volatile / circular / 样式是否可得 |
| **许可证** | 项目约定：**避免 GPL/AGPL 类 copyleft 进入核心依赖路径** |
| **活跃度** | 最后一次发版时间 |
| **建议放置层** | 待 B1 确认的初稿（L0 采集 / L3 校验 / 仅参考） |

---

## 3. 结论摘要（TL;DR）

1. **没有现成库能当核心。** 现有项目全部是 **计算引擎**（回答"算出什么"），本项目需要的是 **无损提取器**（回答"表里有什么、在哪、什么格式"）。引擎吞掉源信息，提取器必须保留源信息。
2. **最强的 `formulas` 也是最危险的依赖**：覆盖最高（90.1%）、最活跃，但许可证为 **EUPL 1.1+（copyleft）**。
3. **`pycel` 直接出局**：GPLv3 + 2021 年后停更。
4. **核心采集层自研（openpyxl）已落地**，本文确认该选择，并给出可选增强（stdlib 包扫描、tokenizer 替换正则）与校验路径（L3 oracle）。
5. **当前实现尚未吃满 openpyxl 的保真潜力**（见 §4.6 缺口清单）——这是"保真"能否成立的真正瓶颈，不是选库问题。

---

## 4. 逐一评估

> 规则：§4.1–§4.8 每个工具必须给出**三行结论**（放置层 / 许可证 / 最后发版）；§4.9 为降级参考项，不适用该规则。

### 4.1 `formulas`（vinci1it2000）

- **定位:** Excel 公式解释器 + 工作簿编译器
- **版本 / 活跃度:** 1.3.4，2026-03-11（活跃）
- **许可证:** **EUPL 1.1+**（copyleft，与 GPL 兼容）
- **能力:** `ExcelModel().loads().finish().calculate()` 编译整个工作簿为 Python 执行图（schedula dispatch）；支持循环引用 `circular=True`；`from_ranges()` 子模型抽取；`compile(inputs, outputs)`；模型可导出 JSON；CLI `build/calc/test/serve`，支持 batch 多场景
- **函数覆盖:** **483 / 536 = 90.1%**（DATE&TIME / ENGINEERING / FINANCIAL / STATISTICAL / LOGICAL / OPERATORS 均 100%；LOOKUP 82.5%、MATH&TRIG 88.8%、TEXT 88.0%；**CUBE / DATABASE / WEB / AUTOMATION 为 0%**）
- **局限:** 产物不是可读 Python 源码（内部 dispatch 图）；不产出坐标级保真清单；缺 CUBE/DATABASE/WEB

> 1. **放置层：** L3 校验 oracle（可选依赖，隔离进程/可选 extra）
> 2. **许可证进核心：** **否**（EUPL 1.1+ copyleft）
> 3. **最后发版：** 2026-03-11

### 4.2 `pycel`

- **版本:** 1.0b30，**2021-10-13**（停更 4 年+）
- **许可证:** **GPLv3**
- **能力:** 生成 graph-based Python（缓存 + 惰性求值）；支持数组 CSE、INDIRECT/OFFSET/INDEX、结构化引用、迭代计算
- **局限:** GPLv3；停更；不编译 VBA；OFFSET 依赖单元格是否已编译

> 1. **放置层：** 仅研究参考（tokenizer / 地址处理思路，**禁止引入代码**）
> 2. **许可证进核心：** **否**（GPLv3）
> 3. **最后发版：** 2021-10-13

### 4.3 `xlcalculator`（koala2 现代化版）

- **版本:** 0.5.0，**2023-02-06**；**许可证: MIT**（0.2.3 起由 GPL-3 改为 MIT）
- **能力:** 工作簿 → Python 状态（可存/可取）；子模型聚焦；求值单格/命名区域/区域/共享公式；可注册自定义函数
- **局限:** **0.1.0 起重构为 AST 求值，不再生成 Python 代码**；不支持数组/CSE 公式；**缺 INDEX / OFFSET / INDIRECT / HLOOKUP / COLUMN / ROW**；VLOOKUP 仅精确匹配；LN / YEARFRAC 与 Excel 有偏差
- **与 plan 差异：** plan 允许其作为插件内实验；本文因上述能力缺口，**不建议投入工程时间**

> 1. **放置层：** 研究参考（AST 设计与精度处理章节）
> 2. **许可证进核心：** 允许（MIT），但能力不足，不建议
> 3. **最后发版：** 2023-02-06

### 4.4 `koala2`

`xlcalculator` 前身，已废弃。不使用。

> 1. **放置层：** 不使用
> 2. **许可证进核心：** 不适用（已废弃）
> 3. **最后发版：** 已停止维护

### 4.5 `FlyingKoala`

- **定位:** xlwings 辅助函数集，Excel 公式可被 Python 调用与单元测试；依赖 xlwings + koala + pandas
- **价值:** "用 Excel 自身做计算引擎来单元测试公式"与增量迁移的方法论
- **局限:** 依赖 Excel COM，**无法在 CI（ubuntu-only）中运行**

> 1. **放置层：** L3 校验思路参考（回归方法论），**不进 CI**
> 2. **许可证进核心：** 否（依赖 Excel COM）
> 3. **最后发版：** N/A（无近期发版；仓库低活跃，需人工确认具体日期）

### 4.6 `openpyxl`（**L0 核心**）

- **许可证:** MIT；**版本:** 已由 `pyproject.toml` 锁定 `openpyxl>=3.1`
- **决定性能力:**
  - `data_only=False` → **公式原文**；`data_only=True` → **缓存值**
  - `number_format`、`data_type`、`comment`、`hyperlink`、merged、tables、defined names、data validation、conditional formatting、sheet 保护/可见性
- **局限:** 不计算；大文件慢（`read_only=True` 可流式加速，但会丢失部分格式/注释信息，与保真目标存在取舍）
- **⚠ 当前实现尚未吃满其保真潜力（缺口清单）:**
  1. **未做 `data_only=True` 二次加载**，`CellInventory`（`schemas/artifacts.py:95-105`）**无 `cached_value` 字段**
  2. 引用解析用正则（`graph/builder.py:20` `REF_RE`），不识别结构化引用 `Table[Col]`，未构造 defined-name 节点
  3. 未采集 fills / fonts / borders / print areas（plan 桶 6 要求）
  4. 未识别 volatile / circular（plan 桶 5 要求）

> 1. **放置层：** **L0 核心采集层（已落地）**
> 2. **许可证进核心：** **是**（MIT）
> 3. **最后发版：** 持续维护

### 4.6b Python 标准库 `zipfile` + `xml.etree`（OOXML 包扫描）

- **许可证:** PSF / PSF-2.0；**零新增依赖**
- **作用:** 识别 openpyxl 无法建模的 workbook 部件：`xl/vbaProject.bin`、`charts/`、`pivotTables/`、`externalLinks/`、`connections.xml`、`drawings/`、`embeddings/`、`media/`，产出 `UnsupportedFeature` — **这是项目"禁止静默丢弃"约定的唯一实现机制**
- **已落地:** `src/excel_to_act/ingest/ooxml_package.py`（`OPAQUE_MARKERS`）
- **⚠ 已知缺陷:** `ooxml_package.py:51` 的 `zipfile.ZipFile(...)` 无异常保护，加密/损坏工作簿会抛出未捕获异常而非产出 `UnsupportedFeature`

> 1. **放置层：** **L0 核心采集层（与 openpyxl 并列，已落地）**
> 2. **许可证进核心：** **是**（标准库）
> 3. **最后发版：** 随 CPython 发布

### 4.7 `fastexcel` / `python-calamine`

- **定位:** Rust calamine 绑定，高速只读；**许可证:** MIT 系（引入前最终核实）
- **⚠ 约束（必须遵守）:** 仅作**旁路加速/预筛**，**不得作为唯一 L0 数据源**。必须实现 `WorkbookReader` 协议，产物须标注 `fidelity='values_only'`，并对未覆盖维度产出 `unsupported_or_opaque` 记录；否则直接与覆盖不变量冲突

> 1. **放置层：** L0 可选旁路加速（非数据源）
> 2. **许可证进核心：** 待核实
> 3. **最后发版：** 持续维护

### 4.8 LibreOffice（headless / soffice）

- **许可证:** MPL 2.0；**用法:** 进程外调用重算与转换，不链接
- **价值:** L3 第二 oracle；价值与局限并存（需安装、启动慢、与 Excel 数值存在已知差异）

> 1. **放置层：** L3 校验 oracle（进程外）
> 2. **许可证进核心：** 不适用（进程外调用，不链接）
> 3. **最后发版：** 持续维护

### 4.8b `xlwings`（plan 点名的 oracle runner）

- **许可证:** BSD-3-Clause
- **价值:** 用 Excel 自身重算，作为数值基线
- **局限:** 依赖 Excel COM；CI 为 `ubuntu-latest`、Python 3.11/3.12（`.github/workflows/ci.yml:19`），**只能在开发者 Windows 机本地跑，不进 CI**

> 1. **放置层：** L3 oracle（本地基线，不进 CI）
> 2. **许可证进核心：** 否
> 3. **最后发版：** 活跃（具体版本待核实）

### 4.9 其他研究参考（不适用三行结论规则）

| 项目 | 定位 | 许可证 | 结论 |
|---|---|---|---|
| `Gridmonger` | JVM/Kotlin Excel 逆向工程可视化 | GPLv3 | 仅算法/交互参考 |
| `orcus` / `ixion` | C++ 表格模型导入与公式引擎 | MPL 2.0 | 仅语法/引擎参考 |
| `PyXLL` / `Excel-DNA` | Excel 内嵌 Python | 商业/混合 | 方向相反，不适用 |
| `SpreadsheetConverter` | Excel → C#/Java | 商业 | 不采用 |
| `Mito` | Jupyter 录制生成 pandas 代码 | 开源 | 是"录制操作"非"解析既有模型"，不适用 |

### 4.10 输入形态：加密 / 受保护 / 二进制与旧格式

| 形态 | 候选工具 | 说明 |
|---|---|---|
| 加密 / 受保护工作簿 | `msoffcrypto-tool` | 许可证与 API 稳定性待核实 |
| `.xlsb` | `pyxlsb` | 公式/格式覆盖度待核实 |
| `.xls` / 旧格式 | LibreOffice 转换 | 进程外转换后再走标准流程 |

**硬约束（项目约定）：** 任何读取失败必须产出 `UnsupportedFeature(severity=error)`，**不得抛出未捕获异常**。当前 `ooxml_package.py:51` 违反此约束。

---

## 5. 横向对比矩阵

| 工具 | 产出可读代码 | 函数覆盖 | 许可证 | 最后发版 | 求公式值 | 结构保真 |
|---|---|---|---|---|---|---|
| `formulas` | 否（dispatch 图 + JSON） | **90.1%** | EUPL 1.1+ | 2026-03 | 是 | 否 |
| `pycel` | 是（graph Python） | 未统计（随需求） | GPLv3 | 2021-10 | 是 | 否 |
| `xlcalculator` | 否（AST 求值） | 中低 | MIT | 2023-02 | 是 | 否 |
| `koala2` | — | 低 | — | 已废弃 | 是 | 否 |
| `openpyxl` | 不适用 | 不适用 | **MIT** | 活跃 | 否（**但可读 Excel 缓存值**） | **是** |
| stdlib `zipfile`+XML | 不适用 | 不适用 | PSF | 随 CPython | 否 | **是（部件级）** |
| `fastexcel` | 不适用 | 不适用 | MIT（待核实） | 活跃 | 否 | 部分（values-only） |
| LibreOffice | 不适用 | 高 | MPL 2.0（进程外） | 活跃 | 是 | 部分 |

**关键对角线：** 唯一具备保真**潜力**的是 `openpyxl`，而它恰恰不计算。且"潜力"能否兑现取决于我们的实现完整度（见 §4.6 缺口清单）。

---

## 6. 许可证风险汇总

| 许可证 | 工具 | 进核心 | 说明 |
|---|---|---|---|
| MIT | openpyxl、xlcalculator、fastexcel | **允许** | 无 copyleft 传染 |
| PSF | stdlib zipfile/xml | **允许** | 标准库 |
| EUPL 1.1+ | formulas | **否** | copyleft；内部自用风险较低，随产物分发需法律确认 |
| GPLv3 | pycel、Gridmonger | **否** | 与项目约定冲突 |
| MPL 2.0 | LibreOffice、orcus/ixion | 不适用 | 进程外调用或仅参考，不链接 |
| BSD-3 | xlwings | 否 | 依赖 Excel COM |

> **加严声明：** plan 原文为 "Avoid **GPL/AGPL** core dependencies"。本文把 **EUPL 1.1+** 也判为禁止进核心，属**加严解释**，须由 F1 ADR 确认，并回写 `docs/plans/phase1_excel_decomposition_plan.md` 相应段落。
> **传递依赖：** 上表仅覆盖直接依赖。`formulas` 传递依赖 `schedula` 等、`xlcalculator` 传递依赖 numpy/openpyxl，**须一并核实**（见 §8）。

---

## 6.5 公式引用解析（tokenizer / AST）专项

plan 明确点名 `formulas` / `xlcalculator` 可用于 **formula reference parsing / evaluation** 实验。本文分别回答：

| 用途 | 结论 |
|---|---|
| **reference parsing（引用解析）** | 首选 **`openpyxl.formula.tokenizer`**（已是核心依赖，MIT，**零新增依赖、零许可证风险**）。**不得为引用解析引入 EUPL/GPL 依赖** |
| **evaluation（求值）** | `formulas` 仅作 L3 oracle（EUPL，隔离）；pycel tokenizer 思路仅参考（GPLv3，**禁止引入代码**） |

当前实现为正则（`graph/builder.py:20`），存在字符串常量内伪地址误判、不识别 `Table[Col]`、未构造 defined-name 节点三类问题。

---

## 7. 放置建议（**非规范性初稿，待 B1 确认**）

| 层 | 采用 | 现有模块 | 缺失模块 |
|---|---|---|---|
| **L0 采集** | `openpyxl` + stdlib `zipfile`/XML（核心）；`fastexcel` 仅旁路加速 | `ingest/openpyxl_reader.py`、`ingest/ooxml_package.py`、`inventory/extractor.py` | `inventory/layout.py`、`inventory/opaque.py`、`ingest/calamine_reader.py` |
| **L1 视图** | **自研，无现成可用** | 无（仅有 `schemas/artifacts.py` 的 `SourceLocation`） | 建议新建 `view/` + `ViewSlice` 契约 |
| **L2 语义推理** | 自研 | `classify/rules.py`、`classify/classifier.py`、`confirm/templates.py` | 无真正语义推理模块；`ActuarialHint`（6 值）与 plan 的 10 个 hint 不一致 |
| **L3 校验 oracle** | `formulas`（可选依赖、隔离）+ LibreOffice（进程外）+ **缓存值对比（零依赖首选）** | **完全缺失** | `validation/{cached_value,formulas_oracle,libreoffice_oracle}.py` + `OracleRunner` 协议 + `ValidationReport` |
| 层外（B1 需补位） | — | `store/local_store.py`、`orchestrator/phase1.py`、`interfaces/cli.py` 已落地 | `report/markdown.py` 缺失（`jinja2` 已声明但未使用） |

**⚠ 缓存值对比的前置条件：** 当前仓库**无 `data_only=True` 加载、无 `cached_value` 字段**，该 oracle **暂不可实现**。必须先补齐缓存值采集（且覆盖计数不得重复计数），否则 L3 只能依赖 `formulas` / LibreOffice。

**⚠ extras 现状：** `pyproject.toml` 已声明 `formula = ["formulas>=1.3", "xlcalculator>=0.5"]`——把 EUPL 包与 MIT 包捆绑在同一个 extra，与本文"隔离"建议相悖。建议拆为 `oracle-formulas`（EUPL，显式 opt-in）与 `xlcalc`（MIT）。

---

## 8. 待核实项（Open Questions）

1. `fastexcel` / `python-calamine` 的确切许可证与公式/格式读取能力 → 引入前实测
2. `formulas` 在 EUPL 下"内部使用 vs 随产物分发"的边界 → 需法律确认（→ F1）
3. `formulas` 对含 VBA / 数据透视表工作簿的失败模式 → 需 fixture 实测
4. **Python 版本可行性**：`requires-python = ">=3.11"` 无上限，CI 仅测 3.11/3.12。须在 3.11/3.12 实测 `pip install formulas / xlcalculator / fastexcel` 可安装性与导入是否成功
5. **openpyxl 自带 tokenizer**（`openpyxl.formula.tokenizer`，openpyxl≥3.1）能否覆盖 A1/跨表/结构化引用/defined name → 若是，作为引用解析首选，替换正则
6. **传递依赖许可证**：`formulas`→`schedula` 等、`xlcalculator`→numpy 等
7. **缓存值缺失时**（工作簿从未由 Excel 重算存盘）该 oracle 如何降级为 warning 而非 error
8. `msoffcrypto-tool` 许可证与 API 稳定性；`pyxlsb` 对 `.xlsb` 的公式/格式覆盖度
9. `openpyxl` `read_only=True` 模式下具体丢失哪些保真维度 → 实测

---

## 9. 验收自检

| 验收项 | 状态 | 证据 | 阻塞项 |
|---|---|---|---|
| 每个工具（§4.1–§4.8）有三行结论 | **完成** | 4.1–4.8b 各含三行结论块 | — |
| 计算引擎类主要候选已覆盖 | **完成**（13 个工具 / 11 节） | §4 全节；§4.9 为降级参考项 | 未纳入清单：SheetJS、pandas、xlrd（均非"Excel→代码/求值"定位） |
| 与 A2/A3/B1/B2/B3/C1 边界无重叠 | **部分完成** | §1 Out of scope 已声明非规范性 | 六份关联文档均未创建，**待其创建后复核** |
| 与仓库现状对齐（引用 `src/` 文件） | **完成** | §0 现状表、§4.6 缺口清单、§7 现有/缺失列 | — |
| 放置建议映射到具体文件 | **完成** | §7 表格「现有模块 / 缺失模块」两列 | — |
| 已转化为可执行动作（PR / extras 调整） | **未完成** | — | 需在 `docs/plans/pr_plan_phase1.md` 新增 PR-13（L3 oracle + `OracleRunner` + cached_value），并调整 `pyproject.toml` extras |
| 版本/许可证数据可核验 | **部分完成** | PyPI 元数据页 + GitHub README | §8.4/§8.6 待实测；`FlyingKoala` 缺确切发版日期 |

---

## 附录：代码侧偏差（不在本文决策范围，转 B3 / C 类文档 / 独立 issue）

1. `inventory/extractor.py:100` 覆盖不变量为**恒等式**（`discovered = recognized + opaque`），等式永远成立，检测不了静默丢失；真正的 `discovered` 须由 `zipfile` 独立枚举
2. `CellInventory` 无 `cached_value` 字段，且 `inventory/README.md` 声称含 "cached values"，与代码漂移
3. `ActuarialHint` 枚举（6 值）与 plan 第 88–97 行的 10 个 hint 不一致
4. `ooxml_package.py:51` `zipfile.ZipFile` 无异常保护，加密/损坏文件会抛未捕获异常
5. `tests/` 未按模块拆分（仅 2 个文件），`examples/` 无静态 fixture
