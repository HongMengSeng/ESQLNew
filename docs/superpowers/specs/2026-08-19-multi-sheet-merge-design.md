# 多 Sheet 勾选与合并导入设计

日期:2026-08-19
状态:已批准(方案 A)

## 背景与问题

一个 Excel 文件通常包含多个 sheet,它们是**不同数据批次**(例如 `损失`、`损失 (UW22KL1)`、`E27FK1`),需要合并去重后导入同一个目标数据库表。

当前程序存在设计混淆:表名输入框同时被当作"源 sheet 名"和"目标数据库表名"使用;且 `ExcelStreamReader.ReadRows(path)` 不带 sheet 参数,默认只读第一个 sheet。因此多 sheet 文件只能导入第一个 sheet。

目标:让用户能全选/多选/单选文件中的 sheet,合并(按表头名对齐、可选去重)后导入同一个目标表。

## 需求

1. 选择 Excel 文件后,弹出**勾选列表**对话框,列出所有 sheet,支持全选 / 全不选 / 多选 / 单选。
2. 勾选对话框中可选择一个**可选的去重键字段**(下拉框,含"不去重"选项)。
3. 目标数据库表名继续用现有表名框选择,不再兼职源 sheet 名。
4. 匹配预览只显示第一个选中 sheet 的列映射与样例行。
5. 导入时流式合并所有选中 sheet,按**表头名**对齐(列序可能不同),若选去重键则按该字段值去重,批量插入目标表。

## 架构

### 组件

**新增:SheetSelectDialog(勾选弹窗)**
- 继承 Form;构造传入 `IList<string> sheetNames` 和第一个选中 sheet 的表头名列表 `IList<string> headers`。
- 顶部 `CheckedListBox` 列出全部 sheet(默认全选)。
- 中部 `ComboBox` 选择去重键字段,首项为"不去重",其余为第一个选中 sheet 的 **Excel 表头名**(不依赖数据库,弹窗时即可用)。
- 底部按钮:全选 / 全不选 / 确定 / 取消。
- 确定后公开属性:
  - `IList<string> SelectedSheets`
  - `string DedupKey`(为空表示不去重;以 Excel 表头名表示)
- 无注释、net462、中文 UI、`FixedToolWindow` + `ControlBox=false`(与现有 ProgressDialog 风格一致,但需要按钮所以保留标题栏关闭)。

**改造:ImportEngine**
- 新增入口 `RunMultiSheet(connStr, table, excelPath, IList<string> sheetNames, string dedupKey, int batchSize, int commitEvery, Action<ImportProgress> onProgress, CancellationToken ct)`。
- 逻辑:
  1. 依次读取每个选中 sheet 的表头;以第一个选中 sheet 的表头为基准,按**表头名**把各 sheet 列对齐到目标表字段(复用 `ColumnMapper.Map` 的字段解析;列序不同时按表头名归一)。
  2. 流式遍历所有选中 sheet 的行,合并成一个统一的"已映射到目标字段"的行流。
  3. 若指定了去重键,按该字段值用 `HashSet<string>` 去重(首现保留)。
  4. 复用 `BatchInserter.Execute` 批量插入。
- 保留现有单 sheet `Run` 入口不变;`Run` 可继续用于单 sheet 场景(或由调用方改为传单元素列表走 `RunMultiSheet`)。

**改造:Preview_Click**
- 预览读取第一个选中 sheet 的表头/样例(给 `ReadHeaders`/`ReadRows` 传该 sheet 名),其余逻辑不变。

**改造:导入流程**
- 调用 `RunMultiSheet`,传入勾选结果与去重键。

**改造:选择文件与新增按钮**
- 选择文件后不再把第一个 sheet 名写入表名框;仅记录文件路径。
- 导入 Tab 新增"选择工作表"按钮:点击弹出 `SheetSelectDialog`,成功后显示"已选 N 个工作表"。
- 表名框专用于目标数据库表名。

## 数据流

```
选文件 → 记录文件路径
  → [选择工作表] → SheetSelectDialog(勾选 sheet + 去重键)
      → 显示已选数
  → [匹配预览] → 读第一个选中 sheet → 列映射 + 样例
  → [导入] → RunMultiSheet(流式合并全部选中 sheet → 按表头名对齐 → 可选去重 → 批量插入)
```

## 错误处理

- 未选择文件 → 提示。
- 未选择工作表或未勾选任何 sheet → 提示。
- 选中 sheet 读取失败 → 报错并指明 sheet 名。
- 选中 sheet 无匹配列 → 沿用现有报错。
- 去重键在选中 sheet 字段中不存在 → 提示并中止(防误删)。

## 测试

- 多 sheet 按表头名对齐:列序不同的两个 sheet 合并后字段正确。
- 去重:指定去重键时合并行数 = 各 sheet 并集;不去重时 = 各 sheet 行数之和。
- 单 sheet 退化:勾选窗单选项,行为与现状一致。
- 现有 55 个测试保持通过。

## 兼容性

- 单个 sheet 文件:勾选窗只有一个选项,全选后导入行为等同现在。
- 现有单 sheet 导入路径 `Run` 保留,不破坏已有调用。
