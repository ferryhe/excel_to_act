# Issue Backlog

由 A1 调研（`docs/research/excel_tooling_survey.md`）与独立审核导出。每个 issue 统一三段式：**做什么 / 提交物 / 怎么检查**。
本文件是 issue 的**源**，同步到 GitHub Issues（仓库 `ferryhe/excel_to_act`）后仍保留在仓库中作为留档。

---

<!-- ISSUE: docs: Phase 1.5 文档与缺口补齐总览（Epic） -->

## 做什么

串联 Phase 1 之后（代号 Phase 1.5）的全部资料与代码缺口工作，**只做总览与索引，不做具体设计**。

背景：A1 调研已确认「现成 Excel→代码 库都不能作核心」，需要自研；同时独立审核发现现有实现存在 5 处保真/健壮性偏差。为避免零散推进，建立统一清单。

**不含：** 任何具体设计决策（归各 A/B/C/D/E/F 子文档）、任何代码实现。

## 提交物

- 本 Epic 仅维护下方清单；子项以独立 issue 存在
- 目录约定（`docs/` 下新建）：`research/`、`design/`、`adr/`、`experiments/`（原有 `plans/`）

## 怎么检查

- [x] 下方每个子 issue 均已创建并可跳转（GitHub Issues #3–#14）
- [x] 子 issue 编号回填到本清单

## 子项清单

| Issue | 类别 | 内容 | 优先级 |
|---|---|---|---|
| **#4** | feat | 补齐 `cached_value` 采集（双次加载） | P0 |
| **#5** | bug | 覆盖不变量恒等式 → `discovered` 独立枚举 | P0 |
| **#6** | feat | 引用解析改用 `openpyxl.formula.tokenizer` | P0 |
| **#7** | feat | L3 oracle 校验层 | P1 |
| **#8** | bug | 加密/损坏/.xlsb/.xls 输入健壮性 | P1 |
| **#9** | chore | pyproject extras 拆分 + 许可证合规检查 | P1 |
| **#10** | docs | A2 Docling/markitdown 保真度评估 | P0 |
| **#11** | docs | B1 L0–L3 分层架构与层间契约 | P0 |
| **#12** | docs | B2 Agent 读取契约 | P0 |
| **#13** | docs | B3 数字与公式保真规则 | P0 |
| **#14** | docs | C1 L1 视图编译器设计 | P0 |
| **#15** | feat | 采集模拟运算表 `dataTable`（What-If 数据表） | P0 |
| **#16** | feat | 采集表单控件 `linkedCell`（`ctrlProps` + `vmlDrawing`） | P0 |
| **#17** | feat | 提取 VBA 源码并抽取 VBA↔单元格依赖边 | P0 |
| **#18** | feat | Step 1 收尾：完备性检查 + handoff 契约输出 | P0 |

---

<!-- ISSUE: feat(ingest): 补齐 cached_value 采集（双次加载） -->

## 做什么

让每个单元格同时保留 **公式原文** 与 **Excel 缓存值**。这是后续"数值对账"与"无 Excel 环境下取基准值"的唯一零依赖来源。

现状：`extractor.py:38` 与 `openpyxl_reader.py:37` 均为 `load_workbook(data_only=False)`；`CellInventory`（`schemas/artifacts.py:95-105`）**无 `cached_value` 字段**；而 `inventory/README.md` 声称含 "cached values"，与代码漂移。

**不含：** 求值/重算（那属于 L3 oracle）；缓存值缺失时的对账流程（另立）。

## 提交物

- `src/excel_to_act/schemas/artifacts.py`：`CellInventory` 新增 `cached_value: str | int | float | bool | None = None` 与 `cached_value_available: bool`
- `src/excel_to_act/inventory/extractor.py`：双次加载（`data_only=False` + `data_only=True`），同一坐标合并
- `src/excel_to_act/ingest/cached_values.py`（**新建**）：值侧读取与标量收敛。**偏差说明**：原计划"同步支持 `openpyxl_reader.py`"不执行——manifest 不持有单元格，在那里再加一次全量读取只会多一次 IO 且无处存放；双次加载的单一事实来源放在 `cached_values.py`，由 extractor 调用
- `schemas/workbook_inventory.schema.json` 已重新导出
- `tests/test_cached_value.py`（4 例：无缓存 / 有缓存 / 常量 / 计数不重复）
- 修正 `inventory/README.md` 描述与代码一致

## 怎么检查

- [x] 对含公式的 fixture：`cell.formula` 非空 且 `cell.cached_value` 非空
- [x] 对纯常量单元格：`cached_value is None` 且不报错
- [x] 工作簿从未被 Excel 重算存盘时：`cached_value_available=False`，产出 `UnsupportedFeature(severity=warning)` 而非抛异常
- [x] 覆盖计数不因双次加载而重复计数（同一格仍计 1）
- [x] `pytest` 与 `ruff check .` 通过

**实现备注**：openpyxl 写出的公式单元格是 `<f>…</f><v></v>`，无法直接产出"带缓存值"的样本。测试用 `inject_cached_values()` 回填 `<v>`，即可得到等价 Excel 存盘效果的工作簿。

---

<!-- ISSUE: bug(inventory): 覆盖不变量为恒等式，discovered 需由 OOXML 独立枚举 -->

## 做什么

当前覆盖不变量**自证成立、无法证伪**：`extractor.py:100` 直接把 `discovered` 写成 `recognized + opaque` 之和，等式永远为真，检测不了"静默丢弃"，与项目核心约定（不支持特性必须记为 opaque）相悖。

目标：让 `discovered_workbook_objects` 由 **独立于 openpyxl 的来源**（OOXML 包扫描）计算得出，使该等式具备 falsifiability。

**不含：** 解析 VBA/图表/透视表内容（只枚举不解释）。

## 提交物

- `src/excel_to_act/inventory/extractor.py`（或新建 `inventory/opaque.py`）：`discovered` 改为由 `scan_ooxml_package` 的部件枚举 + sheet 级实际对象数独立计算
- `src/excel_to_act/schemas/artifacts.py`：必要时为 `CoverageSummary` 增加"差异明细"字段
- `tests/test_coverage_invariant.py`：**构造故意漏采的用例，断言等式不成立时报错**（回归测试）
- 报告层输出 per-sheet 覆盖与差异明细（随 `report/markdown.py` 一并验收）

## 怎么检查

- [ ] 人为在 extractor 中跳过某类单元格 → 覆盖校验**必须失败**（证明可证伪）
- [ ] 正常 fixture：`recognized + opaque == discovered` 且两边来源可追溯到不同代码路径
- [ ] 每个 sheet 都能输出 recognized / opaque / discovered 三个数字
- [ ] `pytest` 通过

---

<!-- ISSUE: feat(graph): 用 openpyxl.formula.tokenizer 替代正则 REF_RE -->

## 做什么

现状 `graph/builder.py:20` 用正则 `REF_RE` 解析公式引用，存在三类缺陷：
1. 字符串常量内的伪地址被误判为引用
2. 不识别结构化引用 `Table[Col]`
3. 从不构造 `GraphNodeKind.name`（defined name）节点

目标：改用 `openpyxl.formula.tokenizer`（openpyxl≥3.1 自带，**MIT、零新增依赖**）做词法解析；语法层（AST）如 tokenizer 不足再自研最小解析器。**明确禁止为引用解析引入 EUPL/GPL 依赖**（`formulas` 只能作 L3 oracle，`pycel` 只能参考思路、禁止引入代码）。

**不含：** 求值；函数语义分析。

## 提交物

- 调研结论（写入 `docs/research/excel_tooling_survey.md` §8.5）：tokenizer 对 A1 / 跨表 / 结构化引用 / defined name 的覆盖结论
- `src/excel_to_act/graph/builder.py`：替换正则实现
- `src/excel_to_act/schemas/artifacts.py`：新增 defined-name 节点与结构化引用节点类型（如需）
- `tests/test_graph_builder.py`：覆盖字符串常量、跨表、结构化引用、defined name、外部引用五类用例
- 解析失败的公式降级为 `UnsupportedFeature`，不得抛异常

## 怎么检查

- [ ] `="A1"` 这类字符串常量**不产生**引用边
- [ ] `=Table1[Col]` 产生结构化引用节点，而非误判为地址
- [ ] `=SUM(MortRate)` 中 defined name 产生 `name` 节点
- [ ] 跨表 `=Assumptions!B7` 正确解析 sheet 名
- [ ] 无法解析的公式产出 `UnsupportedFeature`，不崩溃
- [ ] `pytest` 通过

---

<!-- ISSUE: feat(validation): 新增 L3 oracle 校验层 -->

## 做什么

建立数值回归校验能力：用独立计算源重算工作簿，与自研解析结果对账。

三个 oracle（按优先级）：
1. **缓存值对比**（零依赖首选）——**前置条件：需先完成 cached_value issue**
2. `formulas`（EUPL，可选依赖、隔离进程）
3. LibreOffice headless（MPL 2.0，进程外）
4. `xlwings`（BSD-3，依赖 Excel COM，**仅本地，不进 CI**）

**不含：** 精算语义正确性判断；模型代码生成。

## 提交物

- `src/excel_to_act/validation/cached_value.py`、`formulas_oracle.py`、`libreoffice_oracle.py`
- `src/excel_to_act/plugins/contracts.py`：新增 `OracleRunner` Protocol
- `src/excel_to_act/schemas/artifacts.py`：新增 `ValidationReport`（含容差、差异明细、oracle 来源）
- `docs/plans/pr_plan_phase1.md`：新增 PR-13
- `tests/test_validation_oracle.py`

## 怎么检查

- [ ] 同一 fixture 至少两个 oracle 产出可横向对比的结果
- [ ] 差异超容差 → 产出 `UnsupportedFeature` 而非静默通过
- [ ] 未安装 `formulas` / LibreOffice 时降级为 warning，CI 仍绿
- [ ] `ValidationReport` 可序列化并进入 `store`
- [ ] `pytest` 通过（CI 为 ubuntu，不得依赖 Excel COM）

---

<!-- ISSUE: bug(ingest): 加密/损坏/.xlsb/.xls 输入抛出未捕获异常 -->

## 做什么

现状 `ingest/ooxml_package.py:51` 的 `zipfile.ZipFile(workbook_path)` 无异常保护：加密工作簿、损坏文件会抛出未捕获异常，直接违反项目"禁止静默失败、必须记录"约定。`.xlsb`/`.xls` 目前只走到 `file_type=error`。

目标：任何读取失败都产出 `UnsupportedFeature(severity=error)`；并评估加密/二进制/旧格式的接入路径（`msoffcrypto-tool` / `pyxlsb` / LibreOffice 转换），结论写入 A1 §4.10。

**不含：** 真正实现解密与 .xlsb 解析（先给结论与降级路径）。

## 提交物

- `src/excel_to_act/ingest/ooxml_package.py`：`try/except` 包裹，产出 `UnsupportedFeature(severity=error)`
- `docs/research/excel_tooling_survey.md` §4.10：`msoffcrypto-tool` 许可证与 API 稳定性、`pyxlsb` 公式/格式覆盖度结论
- `tests/test_ingest_robustness.py`：加密/损坏/非 xlsx 三类输入用例

## 怎么检查

- [ ] 传入加密 xlsx：不抛出未捕获异常，产出 `severity=error` 的 `UnsupportedFeature`
- [ ] 传入损坏/截断文件：同上
- [ ] 传入 `.xlsb` / `.xls` / `.csv`：产出明确 error 记录，CLI 退出码可控
- [ ] `pytest` 通过

---

<!-- ISSUE: chore(build): 拆分 pyproject extras 并增加许可证合规检查 -->

## 做什么

现状 `pyproject.toml` 的 `formula = ["formulas>=1.3", "xlcalculator>=0.5"]` 把 **EUPL 包** 与 **MIT 包** 捆绑在同一 extra，与"copyleft 隔离"结论相悖；且 extra 名 `formula` 与包名 `formulas` 不一致，易误用。

同时：`requires-python = ">=3.11"` 无上限，CI 核心依赖矩阵现测 3.11/3.12/3.13；仍需确认待引入的可选库在目标 Python 版本上可安装。

**不含：** 具体库的版本升级策略（另有）。

## 提交物

- `pyproject.toml`：拆为 `oracle-formulas`（EUPL，显式 opt-in）与 `xlcalc`（MIT）；`requires-python` 收敛或显式声明支持矩阵
- `.github/workflows/ci.yml`：新增许可证检查步骤（如 `pip-licenses`），对 EUPL/GPL 出现在**核心依赖**时失败
- A1 §8.4 / §8.6：补 3.11/3.12 实测结果与传递依赖许可证结论

## 怎么检查

- [ ] `pip install -e '.[dev]'` 不引入 `formulas`
- [ ] `pip install -e '.[oracle-formulas]'` 才引入 `formulas`
- [ ] CI 在核心依赖出现 EUPL/GPL 时**失败**
- [ ] 3.11 / 3.12 均可安装并导入 `formulas`、`xlcalculator`、`fastexcel`
- [ ] `ruff check .` 与 `pytest` 通过

---

<!-- ISSUE: docs(research): A2 Docling / markitdown 保真度评估 -->

## 做什么

回答一个问题：**文档解析器（Docling / markitdown / pandas 渲染）能否作为 Excel 保真采集底座？**

已掌握的关键证据：`docling/backend/msexcel_backend.py` 使用 `load_workbook(data_only=True)`（公式全丢）、单元格文本用 `str(cell.value)`（number_format 丢失）、用 flood-fill 聚合非空单元格成表（精确坐标语义丢失）、不提取 defined names / tables / 数据验证 / 条件格式 / 透视 / VBA / 外部链接。

**不含：** 计算引擎评估（A1 已完成）；我们自己怎么做（B/C 类）。

## 提交物

- `docs/research/docling_fidelity_assessment.md`
- **源码级证据表**：每条判定对应 docling 后端的具体行为/位置
- **十项保真度逐项判定**：公式 / number_format / 坐标 / named range / Table / 数据验证 / 条件格式 / 透视 / VBA / 外部链接
- 结论：能否进 L0；若不能，可放在哪一层（如仅用于给 LLM 的粗读）

## 怎么检查

- [ ] 十项判定**每一项**都有源码证据，无"其他"兜底
- [ ] 每条判定附可复现方式（版本 + 文件路径）
- [ ] 结论是明确的"能/不能 + 放在哪层"，无模糊表述
- [ ] 与 A1 §1 Out of scope 无内容重叠

---

<!-- ISSUE: docs(design): B1 L0–L3 分层架构与层间契约 -->

## 做什么

正式定义 L0 采集 / L1 视图 / L2 语义推理 / L3 校验四层，以及层间 artifact 契约。同时**给 `store` / `orchestrator` / `report` 明确归位**（当前四层 taxonomy 未覆盖它们）。

背景：该分层目前只出现在 A1 调研文档里，plan / PR / README 均无定义，A1 §7 仅为非规范性初稿。

**不含：** 各层内部实现细节（归 C 类）；排期（归 D 类）。

## 提交物

- `docs/design/layered_architecture.md`
- 分层图 + 每层职责/输入/输出/禁止事项
- 层间 artifact 契约清单（复用 `schemas/artifacts.py` 现有类型，缺的补）
- **文件级映射表**：每层对应哪些已存在 `src/` 文件、哪些待新建
- 明确 `L1 视图` 与 `A3 编码方法` 的职责边界

## 怎么检查

- [ ] 仓库每个现有模块都能唯一归入某一层（含 store/orchestrator/report）
- [ ] 跨层调用只走 `Artifact` 类型，无裸 dict
- [ ] 映射表逐行给出绝对路径，标注已存在/待新建
- [ ] 与 A1 §7 初稿一致或显式说明差异

---

<!-- ISSUE: docs(design): B2 Agent 读取契约 -->

## 做什么

写死"LLM/agent 如何读本项目产物"的硬约束。核心立场：**agent 读表 = 采样，不是解析**；因此 agent 只能读我们自己产出的确定性 artifact 投影，且输出必须可回填。

**不含：** 具体 prompt 工程；模型选型。

## 提交物

- `docs/design/agent_reading_contract.md`
- 硬规则清单，至少含：只能引用 `view_id` + 地址；输出必须带 `source_location`；遇 `opaque` 必须上报、禁止猜测；禁止凭空造数
- 违规示例（正/反例）
- **机器可校验的检查项**（可落成 pytest）

## 怎么检查

- [ ] 每条规则都能转成一个可断言检查（如输出 JSON 缺 `source_location` 即失败）
- [ ] 至少 3 条正例、3 条反例
- [ ] 与 B1 的层间契约无冲突

---

<!-- ISSUE: docs(design): B3 数字与公式保真规则 -->

## 做什么

定义采集阶段"什么必须原样保留、什么允许转换"的规则，防止精度与公式语义在早期不可逆地丢失。

已知风险：Excel 浮点与 Python 不等价（xlcalculator/formulas 均有专门章节）；日期/百分比若提前格式化则不可恢复；数字应保留原始字符串 + `number_format`。

**不含：** 下游代码生成；求值。

## 提交物

- `docs/design/fidelity_rules.md`
- 规则清单：双次加载、原始字符串保留、`number_format` 必存、共享公式展开、日期不得提前格式化
- **每条规则给出违反检测方式**（字段名 + 断言方式）与反例
- 明确"允许转换"的白名单

## 怎么检查

- [ ] 每条规则对应一个可断言的字段检查
- [ ] 至少覆盖：公式原文、缓存值、number_format、data_type、共享公式、日期
- [ ] 与 `CellInventory` 实际字段一致（含本轮新增的 `cached_value`）

---

<!-- ISSUE: docs(design): C1 L1 视图编译器设计 -->

## 做什么

设计 `inventory.json + dependency_graph.json → LLM 友好视图` 的编译器。核心要求：**保留坐标、可反查、确定性、token 可控**。

借鉴 SpreadsheetLLM / SheetCompressor 的倒排索引思路（相同公式/相同格式只存一次 + 地址列表），但**必须保留 source_location**（原方法会丢坐标，我们反向保留）。

**不含：** 编码/压缩算法选型（归 A3）；prompt 与调用策略（归 B2）。

## 提交物

- `docs/design/views_l1_compiler.md`
- 视图 schema：`view_id`、切片策略、倒排编码格式、token 预算模型
- 接口签名（输入/输出类型，复用现有 `Artifact`）
- 反查规则：视图中任一条 → `sheet!A1`
- 确定性要求：同一输入两次编译产出字节一致

## 怎么检查

- [ ] 视图中任一条记录都能反查到 `sheet!A1`
- [ ] 同一 fixture 两次编译产出字节一致
- [ ] 给出 token 预算公式并对一个示例 sheet 给出估算值
- [ ] 字段复用 `CellInventory` / `SourceLocation`，不另造平行契约

---

<!-- ISSUE: feat(inventory): 采集模拟运算表 dataTable（What-If 数据表） -->

## 做什么

采集 Excel 的**模拟运算表**（Data Table / What-If Analysis）：`<dataTable>` 元素位于 worksheet XML 内，含 `ref`、`rowInputCell`、`colInputCell`（一维/二维）。

这是**精算敏感性分析的核心结构**：一张表 = 一个输入变量的完整情景扫描。目前 openpyxl 静默丢弃该元素，且它不在 `OPAQUE_MARKERS` 内，因此**既不采集也不标 opaque，属静默丢失**，直接违反覆盖不变量。

**不含：** 求解（那是求值）；对二维表展开成具体数值（交给 Step 4）。

## 提交物

- `src/excel_to_act/ingest/data_table.py`：直接读 `xl/worksheets/sheet*.xml` 抽取 `<f t="dataTable">`（openpyxl 无 API 且丢公式文本，必须走 XML）
- `src/excel_to_act/schemas/artifacts.py`：`RangeInventory` 复用，kind 取 `data_table`，metadata 含 `row_input_cell` / `col_input_cell` / `two_dimensional` / `corner_cell` / `formula`
- `src/excel_to_act/inventory/extractor.py`：合并进每个 sheet 的 `ranges`，并计入 `recognized`；同时修正角单元格被 `str(DataTableFormula)` 损坏的问题
- `tests/test_data_table.py`：一维、二维、无表、openpyxl 行为留证四类用例
- 更新 README §4 F 组状态（静默丢失 → 已采集）

## 怎么检查

- [x] 一维（仅 `r1`）与二维（`r1` + `r2`）分别被正确识别
- [x] `r1`/`r2` 由 R1C1 转为 `$B$2` 形式并带 `source_location`
- [x] 未使用模拟运算表的工作簿：不多产记录、不报错
- [x] 该对象计入 `recognized`，不再静默丢失
- [x] `pytest` 与 `ruff check .` 通过

**实现备注**：模拟运算表**不是**独立 OOXML 元素，而是挂在角单元格上的 `<f t="dataTable" ref dt2D r1 r2>`。openpyxl 会读成 `DataTableFormula` 对象但**丢弃公式文本**，原 extractor 的 `str(cell.value)` 会把它变成对象 repr——属静默损坏，不只是丢失。故必须 XML 直读。顺带为 `ArrayFormula` 加了同类保护（文本保留，`ref`/spill 仍待补）。

---

<!-- ISSUE: feat(inventory): 采集表单控件 linkedCell（ctrlProps + vmlDrawing） -->

## 做什么

采集**表单控件与单元格的绑定关系**：复选框、数值调节钮、滚动条、下拉框的 `linkedCell`（及 `fListFillRange`）。

老精算模型几乎都靠这些控件做**场景切换/参数调节**，控件本身就是"输入候选"的强信号。当前 `xl/ctrlProps/` 已补入 `OPAQUE_MARKERS`（只保证不丢），但 `vmlDrawing` 里的控件定义与 `linkedCell` 仍未解析。

**不含：** ActiveX 控件（走 OOXML 另一路径，先标记 opaque）；控件外观/位置。

## 提交物

- `src/excel_to_act/ingest/form_controls.py`：解析 `xl/drawings/vmlDrawing*.vml` + `xl/ctrlProps/*.xml`，产出 `控件名 → linkedCell / 类型 / 取值范围`
- `src/excel_to_act/schemas/artifacts.py`：新增 `FormControl`（或复用 `RangeInventory`，kind=`form_control`），含 `control_type`、`linked_cell`、`list_fill_range`
- `src/excel_to_act/inventory/extractor.py`：挂到对应 sheet 并计入 `recognized`
- `tests/test_form_controls.py`
- 更新 README §4 F 组状态

## 怎么检查

- [ ] 构造含 `linkedCell` 的 vmlDrawing fixture，`linked_cell` 解析为 `sheet!A1`
- [ ] `ctrlProps` 缺失时仍能由 vml 得出 linkedCell（反之亦然），不得崩溃
- [ ] 控件计入 `recognized`，且 `linkedCell` 指向的单元格能被 `classify` 识别为 `input_candidate`
- [ ] `pytest` 与 `ruff check .` 通过

---

<!-- ISSUE: feat(ingest): 提取 VBA 源码并抽取 VBA↔单元格依赖边 -->

## 做什么

两步：**先拿到源码，再抽依赖边**。

1. 用 `oletools`(olevba) 从 `xl/vbaProject.bin` 提取 VBA 模块源码（模块名、类型、过程清单）。可选依赖 `vba`，BSD-3。
2. 从源码中抽取 **VBA ↔ 单元格依赖边**。

**为什么必须做第 2 步**：精算老模型普遍用 `Range("B7")` / `Names("Mort_qx")` 驱动计算，这类边在公式图里完全不可见，只扫公式会得到**断裂的依赖图**（`graph/builder.py` 目前完全不看 VBA）。

**不含：** VBA 语义理解；XLM 宏与 DDE（另立）；执行 VBA。

## 提交物

- `src/excel_to_act/ingest/vba.py`：`extract_vba_project(path)` → 模块源码 + 过程清单；未安装 oletools 时降级为 `UnsupportedFeature(severity=warning)`，**不得 import 失败**
- `src/excel_to_act/inventory/vba_links.py`：`extract_vba_cell_links(source)` → 候选边列表，每条带 `confidence` 与 `unresolved` 标记
- `src/excel_to_act/graph/builder.py`：把边写入 `FormulaGraph`，**`GraphEdge.relationship = "vba_ref"`**
- `src/excel_to_act/schemas/artifacts.py`：为 `GraphEdge` 增加 `confidence` 字段；新增 `GraphNodeKind.vba`；新增 `VbaModule` 模型并挂到 `WorkbookInventory.vba_modules`；**禁止新增平行图契约**
- `src/excel_to_act/orchestrator/phase1.py`：在 graph build 之后调 `extract_vba_project` + `extract_vba_cell_links` + `build_vba_edges`，合并进同一 `FormulaGraph`，并把 `vba_modules` 写回 inventory；未装 oletools 时追加 `vba_extraction_skipped` warning
- `tests/test_vba.py`
- 更新 README §4 G 组状态

## 怎么检查

- [x] 含 VBA 的 `.xlsm`：产出模块源码，过程名可列（逻辑已测；真实二进制样本待补充到 `examples/fixtures`）
- [x] 未装 oletools：`vba_extraction_skipped` warning，CI（仅 `.[dev]`）**仍绿**
- [x] `Range("B7")`、`Range("Inputs!$B$2")`、`Names("Mort_qx")`、`Cells(2,3)` 四类引用被抽出，带 `confidence`（`Range` 带 sheet 0.9 / 不带 0.6，`Names`/`Cells` 低分 + `unresolved`）
- [x] 抽出的边写入 `dependency_graph.json`，`relationship == "vba_ref"`，节点 id 与公式图**同一命名空间**（新增 `vba:` 源节点）
- [x] 无法静态确定的引用（`"B" & i`、经变量间接寻址）标记 `confidence` 低 + `unresolved`，**不假装确定**；`Cells` 不产边
- [x] `pytest` 与 `ruff check .` 通过

**实现备注**：`extract_vba_project` 用 `oletools.olevba.VBA_Parser().extract_macros()`；oletools 为可选依赖，缺失时只降级不崩。字面量匹配只能覆盖 `Range("A1")` / `Names("x")` / `Cells(r,c)` 三类，含字符串拼接或间接寻址的引用仍会留在 `unresolved`，交确认步骤处理。真实 `.xlsm` 端到端样本（需含 `vbaProject.bin`）尚无，待加入 `examples/fixtures` 后补一个集成测试。

---

<!-- ISSUE: feat(verify): Step 1 收尾：完备性检查 + handoff 契约输出 -->

## 做什么

转换完成后，agent **必须**对整个输出做完备性检查，然后产出 handoff 作为 Step 1 的唯一出口。

关键约束：完备性检查**不能复用** `WorkbookInventory.coverage`——其 `discovered` 是 `recognized + opaque` 的恒等式（issue #5），恒为真，证明不了任何事。必须**从包里独立重算对象全集**（worksheet XML 数非空单元格 + 枚举包部件），再与产出比对。

**不含：** 语义正确性检查（那是 Step 3）；跨 run 的回归比对。

## 提交物

- `src/excel_to_act/verify/completeness.py`（新增）：`verify_completeness()` 产出 `CompletenessReport`，含 6 项检查
- `src/excel_to_act/report/handoff.py`（新增）：`build_handoff()` + `render_handoff_markdown()`，产出 `Handoff`
- `src/excel_to_act/schemas/artifacts.py`：`CompletenessCheck` / `CompletenessReport` / `CompletenessStatus` / `HandoffArtifactRef` / `Handoff`；`RunMetadata.completeness_status`
- `src/excel_to_act/store/local_store.py`：`write_text()`、`append_artifacts()`；`completeness.json` / `handoff.json` 纳入校验与索引
- `src/excel_to_act/orchestrator/phase1.py`：graph 之后跑 verify，最后写 handoff（handoff 需等其他产物路径确定）
- `src/excel_to_act/interfaces/cli.py`：打印 completeness 状态与 handoff 路径；`fail` → **退出码 1**
- `schemas/completeness_report.schema.json`、`schemas/handoff.schema.json`
- `tests/test_completeness_handoff.py`
- README §4 新增「Step 1 收尾：完备性检查与 handoff」

## 怎么检查

- [x] `sheets_accounted`：每 sheet 有 `SheetInventory`（error）
- [x] `cells_accounted`：每表单元格数与 worksheet XML 独立扫描一致（error）
- [x] `content_parts_accounted`：部件已采集须**有产出证据**（`xl/tables/` ↔ `table` range），否则 error
- [x] `formulas_linked`：每个公式格有出边或有 unparseable 记录（warning）
- [x] `metadata_parts_accounted`（`docProps/`）与 `coverage_arithmetic` 记为 info，不改变状态
- [x] `status=fail` 时 CLI 退出码 1
- [x] `handoff.json` 列出全部产物（路径 / sha256 / 条数）+ `summary` 计数 + opaque 汇总 + next_actions
- [x] `handoff.md` 是**人读的简短英文摘要**：工作簿速写（表/格/公式/缓存值/名称/表/合并区/dataTable/控件/VBA/图节点边）+ `At a glance` + `Blockers` + `Warnings` + `Artifacts` + `Unresolved` + `Next steps`；run 目录与 `<out>/` 根各一份
- [x] opaque 汇总只统计 `opaque=True` 的项（`missing_cached_values` 这类 warning 不计入）
- [x] `store.read_run()` 能回读并校验 `completeness.json` / `handoff.json`
- [x] `pytest`（24 passed）与 `ruff check .` 通过

**说明**：`coverage_arithmetic` 当前会报 `recognized(25) + opaque(2) != discovered(21)`，这是**预期结果**——它正是 issue #5 暴露出来的真实现象，此前被恒等式掩盖。修 #5 时该检查应从 info 升为 error。
