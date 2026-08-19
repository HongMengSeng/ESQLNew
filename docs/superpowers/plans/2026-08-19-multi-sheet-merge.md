# 多 Sheet 勾选与合并导入实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让用户全选/多选/单选 Excel 文件中的 sheet,合并(按表头名对齐、可选去重)后导入同一个目标数据库表。

**Architecture:** 新增 `SheetSelectDialog` 勾选弹窗(勾选 sheet + 选去重键);`ImportEngine` 新增 `RunMultiSheet` 入口,流式合并多 sheet(按表头名对齐)并可选去重,复用 `BatchInserter`;目标表名继续用现有表名框。选择文件不再把 sheet 名写入表名框。

**Tech Stack:** .NET Framework 4.6.2 WinForms, ExcelDataReader, MySql.Data, xUnit。

## Global Constraints

- net462, LangVersion 7.3
- 代码不写任何注释
- UI 文案为中文
- 现有 55 个测试保持通过
- 每个 Task 结束跑 `dotnet test ESQLNew.sln`,预期全绿
- 运行中 ESQLNew.exe 会锁 exe 导致 MSB3027/MSB3021 → 先 `Stop-Process -Name ESQLNew -Force`
- 仅修改本计划列出的文件

---

### Task 1: 合并引擎(ImportEngine.RunMultiSheet + 辅助方法)

**Files:**
- Modify: `src/ESQLNew/Import/ImportEngine.cs`(新增 `RunMultiSheet` 和 `MergeSheets` 辅助方法)

**Interfaces:**
- Consumes: `ExcelStreamReader.ReadRows(path, sheetName)`、`ReadHeaders(path, sheetName)`、`ColumnMapper.Map(headers, cols, fieldMap)`、`ColumnMapStore.GetMap(path, table)`、`BatchInserter.Execute`、`ImportResult`、`ImportProgress`、`ColumnMapping`、`ColumnInfo`
- Produces:
  - `public static Task<ImportResult> RunMultiSheet(string connStr, string table, string excelPath, IList<string> sheetNames, string dedupKey, int batchSize, int commitEvery, Action<ImportProgress> onProgress, CancellationToken ct)`
  - `internal static IEnumerable<object[]> MergeSheets(string excelPath, IList<string> sheetNames, IList<int> positions, IList<ColumnMapping> canonicalMappings, IList<ColumnInfo> cols, IDictionary<string, string> fieldMap, string dedupKey)`

#### 逻辑说明

`RunMultiSheet` 与现有 `Run` 结构一致,但:
1. 用 `sheetNames[0]`(第一个选中 sheet)读表头作为**规范列基准**(`canonicalMappings`)。
2. 计算 `positions`(matched 索引),与现有 `Run` 相同。
3. 用 `MergeSheets` 生成合并后的行流:第一个 sheet 走 `RebuildRows(ReadRows(path, sheetNames[0]), positions)`;后续 sheet 按表头名对齐到规范字段,并做可选去重。
4. AutoId 逻辑与现有 `Run` 一致(追加 id)。
5. 复用 `BatchInserter.Execute` 批量插入。

#### 去重逻辑

`dedupKey` 是一个 Excel 表头名(来自第一个选中 sheet)。在 `MergeSheets` 中,找到去重字段在**对齐后行**中的位置 `dedupIdx`(遍历 `positions` 中每个 `p`,`canonicalMappings[positions[p]].ExcelColumn == dedupKey`;找不到则抛 `InvalidOperationException("去重键列不存在: " + dedupKey)`)。合并行流中,每个 sheet 的行在写入前,若 `dedupKey` 非空,则按 `row[dedupIdx]` 的字符串值去重(用 `HashSet<string>`,首现保留,空值也参与去重)。`dedupKey` 为空时不创建 `HashSet`。

注意:`dedupIdx` 是"对齐后行"的索引(0..positions.Count-1),不是 `canonicalMappings` 的索引。因为 `RebuildRows(rows, positions)` 产出的对齐行第 `p` 个元素对应 `canonicalMappings[positions[p]]`,所以必须遍历 `positions` 求 `dedupIdx`。

- [ ] **Step 1: 写失败测试**

在 `tests/ESQLNew.Tests/ExcelStreamReaderTests.cs` 的 `XlsxFixture` 类中新增一个 `public static` 方法 `CreateMultiSheet()`,生成一个**双 sheet** 的 xlsx 文件,列序不同,数据确定:

**SheetA(表头 `维修单号 | 担当`)**:3 行数据
- A1 → (维修单号=1001, 担当=张三)
- A2 → (维修单号=1002, 担当=李四)
- A3 → (维修单号=1003, 担当=王五)

**SheetB(表头 `担当 | 维修单号`,列序相反)**:3 行数据
- B1 → (担当=张三, 维修单号=1001)  ← 与 A1 重复(同 1001)
- B2 → (担当=赵六, 维修单号=1004)
- B3 → (担当=孙七, 维修单号=1005)

`CreateMultiSheet()` 的实现参考现有 `Create()` 的 zip 手写方式,但 `workbook.xml` 定义两个 `<sheet>`(SheetA、SheetB),`workbook.xml.rels` 引用 sheet1.xml、sheet2.xml,`sharedStrings.xml` 包含所有表头与数据字符串。返回文件路径。另加 `public static void DeleteMultiSheet(string path)`(与 `Delete` 同实现,可复用 `Delete`)。

**然后**在 `tests/ESQLNew.Tests/ImportEngineTests.cs` 增加失败测试:

```csharp
[Fact]
public void MergeSheets_AlignsAndDedups()
{
    string path = ESQLNew.Tests.ExcelStreamReaderTests.XlsxFixture.CreateMultiSheet();
    try
    {
        var positions = new List<int> { 0, 1 };
        var canonical = new List<ColumnMapping>
        {
            new ColumnMapping { ExcelColumn = "维修单号", TableField = "repair_order_no", Matched = true },
            new ColumnMapping { ExcelColumn = "担当", TableField = "person_in_charge", Matched = true }
        };
        var cols = new List<ColumnInfo>
        {
            new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
            new ColumnInfo { Name = "person_in_charge", DataType = "varchar", IsNullable = true, MaxLength = 50 }
        };
        var fieldMap = new Dictionary<string, string>
        {
            { "维修单号", "repair_order_no" },
            { "担当", "person_in_charge" }
        };
        var merged = ImportEngine.MergeSheets(path, new List<string> { "SheetA", "SheetB" },
            positions, canonical, cols, fieldMap, "维修单号").ToList();
        // SheetA 3 行 + SheetB 去重后仅剩 1004、1005 两行 => 共 5 行
        Assert.Equal(5, merged.Count);
        // 验证第 0 行对齐为 [维修单号=1001, 担当=张三]
        Assert.Equal("1001", merged[0][0] == null ? "" : merged[0][0].ToString());
        Assert.Equal("张三", merged[0][1] == null ? "" : merged[0][1].ToString());
    }
    finally
    {
        ESQLNew.Tests.ExcelStreamReaderTests.XlsxFixture.DeleteMultiSheet(path);
    }
}
```

**关键**:测试依赖两个引用:
1. `XlsxFixture.CreateMultiSheet` 与 `DeleteMultiSheet` 必须是 `public`(现有 `XlsxFixture` 是 `private` 嵌套类,需改为 `public`,且 `ExcelStreamReaderTests` 需保持可访问)。若不便改 `ExcelStreamReaderTests`,实现者可在 `ImportEngineTests.cs` 内复制一份 `XlsxFixture`(重命名 `ImportFixture`)并实现 `CreateMultiSheet`。**任选其一,保证 `CreateMultiSheet`/`DeleteMultiSheet` 为 public 且可被 ImportEngineTests 调用。**
2. 字段值在合并行中的对齐:第一个 sheet 走 `RebuildRows(rows, positions)`,positions=[0,1] → 行 = [列0=维修单号, 列1=担当]。

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~MergeSheets"`
Expected: 编译失败(找不到 `ImportEngine.MergeSheets`)→ 红。

- [ ] **Step 3: 实现 `MergeSheets` 和 `RunMultiSheet`**

在 `ImportEngine.cs` 添加(参考现有 `RebuildRows`/`Run` 风格):

```csharp
public static Task<ImportResult> RunMultiSheet(string connStr, string table, string excelPath,
    IList<string> sheetNames, string dedupKey, int batchSize, int commitEvery,
    Action<ImportProgress> onProgress, CancellationToken ct)
{
    if (sheetNames == null || sheetNames.Count == 0)
        throw new InvalidOperationException("未选择工作表");

    IList<ColumnInfo> cols;
    try
    {
        cols = GetTableColumns(connStr, table);
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException("无法读取目标表结构,请检查连接与表名:" + ex.Message, ex);
    }
    if (cols.Count == 0)
        throw new InvalidOperationException("目标表不存在或无列:" + table);

    IList<string> headers;
    try
    {
        headers = ExcelStreamReader.ReadHeaders(excelPath, sheetNames[0]);
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException("无法读取 Excel 表头:" + ex.Message, ex);
    }

    var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
    var mappings = ColumnMapper.Map(headers, cols, fieldMap);
    var positions = new List<int>();
    for (int i = 0; i < mappings.Count; i++)
        if (mappings[i].Matched)
            positions.Add(i);

    if (positions.Count == 0)
        throw new InvalidOperationException("Excel 表头与目标表字段无匹配,请检查列名:" + table);

    IEnumerable<object[]> rows = MergeSheets(excelPath, sheetNames, positions, mappings, cols, fieldMap, dedupKey);
    if (AutoIdNeeded(mappings, cols))
    {
        long nextId = GetMaxId(connStr, table) + 1;
        rows = WithAutoId(rows, nextId);
        mappings.Add(new ColumnMapping
        {
            ExcelColumn = null,
            TableField = "id",
            Matched = true,
            TableColumn = GetColumn(cols, "id")
        });
    }

    var result = new ImportResult();
    BatchInserter.Execute(connStr, table, mappings, rows, batchSize, commitEvery, onProgress, result, ct);
    return Task.FromResult(result);
}

internal static IEnumerable<object[]> MergeSheets(string excelPath, IList<string> sheetNames,
    IList<int> positions, IList<ColumnMapping> canonicalMappings, IList<ColumnInfo> cols,
    IDictionary<string, string> fieldMap, string dedupKey)
{
    int dedupIdx = -1;
    if (!string.IsNullOrEmpty(dedupKey))
    {
        // 找去重字段在"对齐后行"中的位置 p:对齐行第 p 个元素对应 canonicalMappings[positions[p]]
        for (int p = 0; p < positions.Count; p++)
        {
            int src = positions[p];
            if (canonicalMappings[src].Matched &&
                string.Equals(canonicalMappings[src].ExcelColumn, dedupKey, StringComparison.OrdinalIgnoreCase))
            { dedupIdx = p; break; }
        }
        if (dedupIdx < 0)
            throw new InvalidOperationException("去重键列不存在: " + dedupKey);
    }

    var seen = dedupIdx >= 0 ? new HashSet<string>() : null;

    bool first = true;
    foreach (var sheet in sheetNames)
    {
        if (first)
        {
            foreach (var row in RebuildRows(ExcelStreamReader.ReadRows(excelPath, sheet), positions))
            {
                if (seen != null)
                {
                    string key = row[dedupIdx] == null ? "" : row[dedupIdx].ToString();
                    if (!seen.Add(key)) continue;
                }
                yield return row;
            }
            first = false;
        }
        else
        {
            IList<string> sheetHeaders = ExcelStreamReader.ReadHeaders(excelPath, sheet);
            var sheetMappings = ColumnMapper.Map(sheetHeaders, cols, fieldMap);
            var fieldToIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < sheetMappings.Count; i++)
                if (sheetMappings[i].Matched)
                    fieldToIdx[sheetMappings[i].TableField] = i;

            foreach (var raw in ExcelStreamReader.ReadRows(excelPath, sheet))
            {
                var row = new object[positions.Count];
                for (int i = 0; i < positions.Count; i++)
                {
                    int idx;
                    if (fieldToIdx.TryGetValue(canonicalMappings[positions[i]].TableField, out idx) && idx < raw.Length)
                        row[i] = raw[idx];
                    else
                        row[i] = null;
                }
                if (seen != null)
                {
                    string key = row[dedupIdx] == null ? "" : row[dedupIdx].ToString();
                    if (!seen.Add(key)) continue;
                }
                yield return row;
            }
        }
    }
}
```

**重要修正:** `canonicalMappings[positions[i]].TableField` 用于定位字段;但注意第一个 sheet 的 `RebuildRows(rows, positions)` 产出的是 `raw[positions[i]]`,即规范顺序。后续 sheet 需对齐到相同规范顺序,因此用 `canonicalMappings[positions[i]].TableField` 在 `fieldToIdx` 中查该字段在后续 sheet 的列索引。

- [ ] **Step 4: 运行测试确认通过**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~MergeSheets"`
Expected: PASS。

- [ ] **Step 5: 运行全部测试**

Run: `dotnet test ESQLNew.sln`
Expected: 56+ 通过(原 55 + 新增)。

- [ ] **Step 6: 提交**

```bash
git add tests/ESQLNew.Tests/ExcelStreamReaderTests.cs tests/ESQLNew.Tests/ImportEngineTests.cs src/ESQLNew/Import/ImportEngine.cs
git commit -m "feat: add multi-sheet merge import engine"
```

---

### Task 2: SheetSelectDialog 勾选弹窗

**Files:**
- Create: `src/ESQLNew/MainForm.SheetSelectDialog.cs`(与 `MainForm.ProgressDialog.cs` 同风格,partial class MainForm 内嵌套类)

**Interfaces:**
- Consumes: 无外部依赖;输入 `IList<string> sheetNames`、`IList<string> headers`(第一个选中 sheet 表头)
- Produces:
  - `private class SheetSelectDialog : Form`
  - 构造:`SheetSelectDialog(IList<string> sheetNames, IList<string> headers)`
  - 属性:`public IList<string> SelectedSheets { get; }`、`public string DedupKey { get; }`(DedupKey 为空表示不去重)

#### UI 布局

- `CheckedListBox`(顶部,列出全部 sheet,默认全选)
- `ComboBox`(中部,去重键:首项"不去重",其余为 `headers`)
- 按钮行:全选 / 全不选 / 确定 / 取消
- `FormBorderStyle.FixedToolWindow`、`ShowInTaskbar=false`、`StartPosition.CenterParent`
- 确定后填充 `SelectedSheets`(勾选的 sheet,顺序保持列表顺序)和 `DedupKey`(去重键下拉选中的值;若选"不去重"则为空)

- [ ] **Step 1: 写失败测试(编译级)**

SheetSelectDialog 是 UI 类,net462 WinForms,不便单测。本任务**不写单元测试**(UI 类,与 ProgressDialog 同理)。改用**编译验证**:确认 `MainForm.SheetSelectDialog.cs` 能编译通过。

Run: `dotnet build ESQLNew.sln -c Release`
Expected: 0 错误(此时文件尚不存在则失败)。先创建空壳类再 build。

- [ ] **Step 2: 实现 SheetSelectDialog**

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class SheetSelectDialog : Form
        {
            private readonly IList<string> _sheetNames;
            private readonly CheckedListBox _list;

            public IList<string> SelectedSheets { get; private set; }
            public string DedupKey { get; private set; }

            public SheetSelectDialog(IList<string> sheetNames, IList<string> headers)
            {
                _sheetNames = sheetNames;
                Text = "选择工作表";
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ShowInTaskbar = false;
                ClientSize = new Size(360, 300);

                _list = new CheckedListBox { Dock = DockStyle.Fill };
                foreach (var n in sheetNames)
                    _list.Items.Add(n, true);

                var dedup = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                dedup.Items.Add("不去重");
                if (headers != null)
                    foreach (var h in headers)
                        if (!string.IsNullOrEmpty(h))
                            dedup.Items.Add(h);
                dedup.SelectedIndex = 0;

                var all = new Button { Text = "全选", Dock = DockStyle.Fill };
                var none = new Button { Text = "全不选", Dock = DockStyle.Fill };
                var ok = new Button { Text = "确定", Dock = DockStyle.Fill, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Dock = DockStyle.Fill, DialogResult = DialogResult.Cancel };

                all.Click += (s, e) =>
                {
                    for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, true);
                };
                none.Click += (s, e) =>
                {
                    for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, false);
                };

                var btnFlow = new FlowLayoutPanel { Dock = DockStyle.Fill };
                btnFlow.Controls.Add(all);
                btnFlow.Controls.Add(none);
                btnFlow.Controls.Add(ok);
                btnFlow.Controls.Add(cancel);

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
                layout.Controls.Add(_list, 0, 0);
                layout.Controls.Add(dedup, 0, 1);
                layout.Controls.Add(btnFlow, 0, 2);
                Controls.Add(layout);

                AcceptButton = ok;
                CancelButton = cancel;

                ok.Click += Ok_Click;
            }

            private void Ok_Click(object sender, EventArgs e)
            {
                var selected = new List<string>();
                for (int i = 0; i < _list.Items.Count; i++)
                    if (_list.GetItemChecked(i))
                        selected.Add((string)_list.Items[i]);
                SelectedSheets = selected;
                DedupKey = "";
            }
        }
    }
}
```

**注意:** 去重键下拉的值 `DedupKey` 需从 ComboBox 读取(非"不去重"时)。上面 `Ok_Click` 只演示了 SelectedSheets;实际实现需把 `ComboBox dedup` 提升为字段,在 `Ok_Click` 中:`DedupKey = dedup.SelectedIndex <= 0 ? "" : (string)dedup.SelectedItem;`。请实现者补全。

- [ ] **Step 3: 编译验证**

Run: `dotnet build ESQLNew.sln -c Release`
Expected: 0 错误。

- [ ] **Step 4: 提交**

```bash
git add src/ESQLNew/MainForm.SheetSelectDialog.cs
git commit -m "feat: add sheet select dialog"
```

---

### Task 3: 接入 UI(选择文件、预览、导入、新增按钮)

**Files:**
- Modify: `src/ESQLNew/MainForm.Tabs.cs`

**Interfaces:**
- Consumes: `SheetSelectDialog`(Task 2)、`ImportEngine.RunMultiSheet`(Task 1)、`ExcelStreamReader.ReadHeaders(path, sheet)`、`ReadRows(path, sheet)`、`ColumnMapper.Map`
- Produces:
  - 新字段:`private Button _sheetSelectButton;`、`private IList<string> _selectedSheets;`、`private string _dedupKey;`、`private string _lastSelectedSheet;`
  - 新方法:`private void ChooseSheet_Click(object sender, EventArgs e)`

#### 改造点

1. **字段声明区**(约 38-47 行)新增:`private Button _sheetSelectButton;` `private IList<string> _selectedSheets;` `private string _dedupKey;`

2. **控件初始化**(约 167-215 行):新增"选择工作表"按钮,放在文件选择行或目标表行。推荐放在"目标表"那一行(`tableFlow`)前或文件行旁。具体:在 `layout` 第 1 行"目标表"标签前新增一行"工作表",或把按钮并到文件行。**实现者按 UI 合理布局**,新增 `_sheetSelectButton = new Button { Text = "选择工作表", Dock = DockStyle.Fill };`,并 `_sheetSelectButton.Click += ChooseSheet_Click;`,加入布局。

3. **ChooseFile_Click**(约 649-672 行):**移除** `_tableComboBox.Text = names[0]`(第 665 行)。改为选择文件后,重置 `_selectedSheets = null; _dedupKey = null;`(等用户点"选择工作表")。保留读取 sheet 名用于校验文件可读性。

4. **新增 ChooseSheet_Click**:
```csharp
private void ChooseSheet_Click(object sender, EventArgs e)
{
    string path = _fileTextBox.Text.Trim();
    if (path.Length == 0)
    {
        MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }
    try
    {
        var names = ExcelStreamReader.SheetNames(path);
        var firstHeaders = names.Count > 0 ? ExcelStreamReader.ReadHeaders(path, names[0]) : new List<string>();
        using (var dlg = new SheetSelectDialog(names, firstHeaders))
        {
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            _selectedSheets = dlg.SelectedSheets;
            _dedupKey = dlg.DedupKey;
            _statusLabel.Text = "已选 " + (_selectedSheets != null ? _selectedSheets.Count : 0) + " 个工作表"
                + (string.IsNullOrEmpty(_dedupKey) ? "" : ",去重键:" + _dedupKey);
        }
    }
    catch (Exception ex)
    {
        _statusLabel.Text = "无法读取工作表:" + ex.Message;
    }
}
```

5. **Preview_Click**(约 674 行):确定当前 sheet。若 `_selectedSheets` 为空(用户没点"选择工作表"),提示;否则取 `_selectedSheets[0]`。把 `ExcelStreamReader.ReadHeaders(path)` 改为 `ReadHeaders(path, currentSheet)`,样例行 `ReadRows` 也传 `currentSheet`。**注意** `ShowSample`(约 921 行)内部用 `ReadRows(_fileTextBox.Text.Trim())` 需改为读当前 sheet——通过新增字段 `_currentPreviewSheet` 或在 Preview_Click 中把当前 sheet 存到字段,供 `ShowSample` 使用。

6. **StartImport_Click**(约 823 行):若 `_selectedSheets == null || _selectedSheets.Count == 0`,提示"请先选择工作表"。调用改为:
```csharp
var result = await Task.Run(() =>
    ImportEngine.RunMultiSheet(connStr, table, path, _selectedSheets, _dedupKey, batch, commit, progressCb, System.Threading.CancellationToken.None).Result);
```

7. **控件禁用/恢复**:把 `_sheetSelectButton` 加入 Preview/Import 的禁用与恢复集合(与 `_chooseFileButton` 等同步)。

#### 测试

本 Task 是 UI 集成,不新增单元测试(UI 逻辑)。验证方式:`dotnet build` 0 错误 + 现有测试全绿 + 手动/自动化跑通单 sheet 与多 sheet 导入。

- [ ] **Step 1: 实现所有 UI 改造**(上述 1-7)

- [ ] **Step 2: 构建 + 测试**

Run: `dotnet build ESQLNew.sln -c Release`(0 错误)→ `dotnet test ESQLNew.sln`(全绿)

- [ ] **Step 3: 真实验证多 sheet 合并导入**

用 `B:\17年4-3月维修单信息(1).xlsx`(3 sheet:损失/损失(UW22KL1)/E27FK1),通过临时 probe(参考项目已有 probe 模式)调用 `ImportEngine.RunMultiSheet` 导入测试表 `sx_qms_repair_info_history1`(需先清空该表)。验证:
- 若选全部 3 sheet 且去重键=维修单号 → 结果行数应约为 77,285(三表并集去重)
- 若只选 `损失` → 83,046 行
- 输出耗时与成功/失败数

- [ ] **Step 4: 提交**

```bash
git add src/ESQLNew/MainForm.Tabs.cs
git commit -m "feat: wire multi-sheet select into preview and import"
```
