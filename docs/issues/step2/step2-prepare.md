## 做什么

交付 `excel-to-act step2 prepare`：把已验证的 Step 1 / Step 2 产物准备为 agent 可以逐步读取的目录包，自动生成英文人工与 AI 交接。后续查询复用准备结果，不为每个问题重新解析原始 Excel。

本项是三个顺序实现项中的第 1 项，建议对应一个 PR。已有 #6、#13、#14、#18–#21 是基础，不重复实现。基线 main 为 `9251019a3b77fe12ce36a811d7392bc781219d0d`；当前本地工作树还包含已审核、未发布的 Step 1 命名空间修复及 checkbox/ActiveX/VBA 交接改动。实施前确认这些前置能力可用，并在 PR 中明确前置提交或 PR，不重新开发同一套解析器。

## CLI 契约

```text
excel-to-act step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]
```

- 输入为原生 `index.json` 和对应 Step 1 root；原生 `step1.v1` / `step1.batch.v1` 仍由现有 index 命令处理。
- `--scope` 是源 SHA/run 绑定的 `analysis.scope.v1` 用户范围决定，不是 AI handoff 或数据快照。不提供时保留全部工作表。
- 检查源身份、索引引用、scope 工作表与决定的一致性；批次保留所有条目，失败条目保留诊断，不能静默丢弃。
- `--dry-run` 返回输入、能力和计划文件，不创建目录、文件或恢复历史。
- `--resume` 依据源引用、scope、编译版本、参数、完整必需产物集合和产物哈希判断复用；普通准备与读取不消耗既有三次 indexing recovery 预算。
- manifest 中 `roots` 明确映射 `reading`、`step1` 和 `step2_index`；文件引用保持相对路径。外部 root 映射属于准备输入，变更时会生成新 revision。
- stdout 返回 JSON，至少包含状态、source/run 身份、产物引用、诊断与 next action；完整性失败返回结构化诊断和非零退出码。
- Step 2 handoff 的 `artifacts` 使用 `root` 与 root-relative `path`；audit/snapshot 另带 SHA-256 和字节数，manifest 自身只引用路径。源行保留每个来源的状态、下一步准备状态和诊断。

## 提交物

1. 可运行、可发现的 prepare CLI；`step2 tools` 的 `reader_commands` 只登记本 PR 已实现的能力。
2. 读取目录包：

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

3. manifest 绑定 source/run、原生 index、scope、编译版本/参数、view 和 graph 哈希，以及供后续 selector 使用的查找信息与允许读取的依赖范围。引用使用明确 roots 和相对路径；原始 SourceLocation 保留原样。
4. 复用现有 `WorkbookView` / `ViewRecord` / `SourceLocation`。准备阶段将经过原生索引哈希、typed model 和 workbook 身份检查的 checkbox/ActiveX artifact 纳入 canonical sheet views；完整 typed record 为 facts，位置由实际 owner/part/shape 字段确定，linked cell 不伪造为控件本身的 A1 位置。新增读取记录不改变 Step 1 覆盖率分母。
5. 从已校验的 inventory、名称声明和支持的静态引用生成 dependency audit / snapshot，复用准备阶段的源数据读取。输出绑定源 SHA、run、scope 哈希，并在 manifest 中登记路径和 SHA-256。旧 scope 中没有哈希的 sidecar 文件名仅为历史提示，不作为权威输入；不导入手工修改的依赖值。
6. 一份已验证上下文同时渲染英文人工和机器交接，区分 source quality、reference integrity、static readiness 与尚未验证的 runtime/numerical behavior；scope 改动生成新的读取修订，不编辑 checksum-bound final 产物。
7. 必要的 runtime models / JSON Schema、README 和 agent 准备阶段说明，以及 CLI 自动回归。将已审核的 `docs/design/progressive_excel_exploration.md` 随首个 PR 纳入仓库，并把本 Issue 的 CLI 契约落实到文档。

## 边界与依赖

- 本项止于准备包与交接：不新增 query / trace，不实现证据包的 agent claim 验证，不做业务解释、宏执行、重算后端或 Python 生成。
- 不替换原生 index / source handoff，不改变 Step 1 fidelity、opaque 或 promotion 的语义，不因准备失败降低原有检查门槛。
- 保留原始 source/inventory 和 VBA 文件引用，不重复复制大型源产物；views 按可分别读取的文件保存，graph 可复用，不为将来的问题建立新数据库或 agent 框架。
- 缓存缺失和动态引用明确报告；静态审计不声称依赖全集已知，缓存不声称代表改参后的结果。
- 第 2 项消费本项 manifest / canonical views；第 3 项消费本项已保存的 graph。本项提供必要的数据契约，不提前实现后续查询和遍历算法。

## 怎么检查

- [ ] CLI 的 help、tool catalogue、返回 JSON 和失败退出码与契约一致。
- [ ] 错误 source/run scope、冲突工作表清单、缺失/变更原生引用均有结构化失败；原始 final 和原生 index 字节不变。
- [ ] absent/existing 输出目录下 dry-run 均无写入；不修改 recovery history。
- [ ] unchanged resume 复用；scope/source/compiler/options 改变时正确失效，已变更、缺失或从 manifest 输出集合遗漏的必需准备产物不能复用。
- [ ] 读取包可通过明确 roots 重新定位；保留原始来源位置，避免只靠当前机器绝对路径。
- [ ] checkbox/ActiveX canonical 记录可以从原模块 artifact 回核，并保留完整 facts / source 身份；bindings 不进入逻辑对象覆盖分母。
- [ ] 未引用的旧 sidecar 被修改不能改变生成的数据；生成的 snapshot/audit 被修改或缺失会让复用/校验失败。
- [ ] Pricing 同一源版本、四个获批排除 tab 下，可从已校验源数据生成 864 个支持的静态入站引用、12 个范围、1,464 条存储 cell 记录（其中 1,364 条有公式），保持空单元格与缓存可用性语义，不猜值、不重算。
- [ ] fixture CI 与本地真实 Pricing CLI 运行均保存可检查证据；真实输入和大型输出保留在本地产物目录，不提交到仓库。
- [ ] 记录本项实际 inventory 解析和 graph 构建次数，证明准备产物可被复用；相关测试、Ruff 和 CI 通过。
