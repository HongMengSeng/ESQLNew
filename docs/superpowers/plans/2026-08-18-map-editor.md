# ESQLNew 界面编辑列映射 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在导入 Tab 加"编辑映射"按钮,弹窗逐列编辑当前目标表的列映射并写回配置文件。

**Architecture:** 新增 `MainForm.ColumnMapDialog.cs`(partial 类),弹窗用 DataGridView(只读 Excel 表头列 + 英文字段下拉组合框列);把"从下拉结果生成映射字典"抽成 `ColumnMapStore` 的纯静态方法以便测试;导入 Tab 加按钮,点击读取表头+字段+当前映射后开弹窗,保存后重跑预览。

**Tech Stack:** .NET Framework 4.6.2 WinForms, Newtonsoft.Json 13.0.3, xUnit 2.4.2。

## Global Constraints

- 目标框架 net462,LangVersion 7.3,代码不写任何注释,UI 中文。
- 构建:`dotnet build ESQLNew.sln` 0 错误;测试:`dotnet test ESQLNew.sln` 全绿(当前 51 个)。
- 配置文件路径:`%APPDATA%\ESQLNew\column-maps.json`(经 `ColumnMapStore.ConfigPath`)。
- 编辑/保存只针对当前目标表,不影响其它表。
- 英文字段下拉候选 = `ImportEngine.GetTableColumns(当前连接, 目标表)` 的字段名。
- 未匹配列默认下拉为空;保存只取下拉非空的行的映射。
- 分支:`main`,不 push(除非用户要求)。
- 弹窗为 UI 不做单元测试;可测的纯逻辑抽到 ColumnMapStore。

---

### Task 1: 映射字典收集辅助方法(可测)

**Files:**
- Modify: `src/ESQLNew/Core/ColumnMapStore.cs`
- Test: `tests/ESQLNew.Tests/ColumnMapStoreTests.cs`

**Interfaces:**
- Produces:
  - `static Dictionary<string, string> BuildMap(IList<string> excelHeaders, IList<string> selectedFields)` — 将两个等长列表打包为"表头→字段"字典,跳过 selectedFields 中为 null/空白 的项;表头 trim 后为空也跳过。假定长度相等(由调用方保证)。
  - 若 excelHeaders 与 selectedFields 长度不等,返回空字典(防御)。

- [ ] **Step 1: 写失败的测试**

`tests/ESQLNew.Tests/ColumnMapStoreTests.cs` 追加:

```csharp
        [Fact]
        public void BuildMap_CombinesHeadersAndFields_SkipsEmpty()
        {
            var headers = new List<string> { "序号", "担当", "维修单号" };
            var fields = new List<string> { "seq_no", "", null };
            var map = ColumnMapStore.BuildMap(headers, fields);
            Assert.Equal(1, map.Count);
            Assert.Equal("seq_no", map["序号"]);
        }

        [Fact]
        public void BuildMap_TrimsHeaderAndSkipsBlankHeader()
        {
            var headers = new List<string> { " 序号 ", "", "  " };
            var fields = new List<string> { "seq_no", "order_no", "x" };
            var map = ColumnMapStore.BuildMap(headers, fields);
            Assert.Equal(1, map.Count);
            Assert.Equal("seq_no", map["序号"]);
        }

        [Fact]
        public void BuildMap_LengthMismatch_ReturnsEmpty()
        {
            var map = ColumnMapStore.BuildMap(new List<string> { "a" }, new List<string> { "x", "y" });
            Assert.Empty(map);
        }
```

- [ ] **Step 2: 跑测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ColumnMapStoreTests"`
Expected: 3 个新测试失败(BuildMap 未定义)。

- [ ] **Step 3: 实现 BuildMap**

在 `src/ESQLNew/Core/ColumnMapStore.cs` 的 `GetMap` 方法之后追加:

```csharp
        public static Dictionary<string, string> BuildMap(IList<string> excelHeaders, IList<string> selectedFields)
        {
            var result = new Dictionary<string, string>();
            if (excelHeaders == null || selectedFields == null || excelHeaders.Count != selectedFields.Count)
                return result;
            for (int i = 0; i < excelHeaders.Count; i++)
            {
                var header = excelHeaders[i] == null ? "" : excelHeaders[i].Trim();
                var field = selectedFields[i] == null ? "" : selectedFields[i].Trim();
                if (header.Length == 0 || field.Length == 0) continue;
                result[header] = field;
            }
            return result;
        }
```

需在文件顶部确认 `using System.Collections.Generic;` 已存在(已有)。

- [ ] **Step 4: 跑测试确认通过**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ColumnMapStoreTests"`
Expected: 全部通过(原 6 + 3 = 9)。

- [ ] **Step 5: 提交**

```bash
git add src/ESQLNew/Core/ColumnMapStore.cs tests/ESQLNew.Tests/ColumnMapStoreTests.cs
git commit -m "feat: add BuildMap helper for map dialog result collection"
```

---

### Task 2: 编辑映射弹窗 + 导入 Tab 按钮接入

**Files:**
- Create: `src/ESQLNew/MainForm.ColumnMapDialog.cs`
- Modify: `src/ESQLNew/MainForm.Tabs.cs`

**Interfaces:**
- Consumes: `ColumnMapStore.GetMap/ConfigPath/Save/Load/BuildMap`(Task 1)、`ImportEngine.GetTableColumns`、`ExcelStreamReader.ReadHeaders`。
- Produces:
  - `partial class MainForm` 的私有嵌套类 `ColumnMapDialog : Form`
    - 构造:`ColumnMapDialog(string table, IList<string> excelHeaders, IList<ColumnInfo> tableColumns, Dictionary<string,string> currentMap)`
    - 属性:`Dictionary<string,string> Result`
  - `MainForm` 私有方法 `void EditMapping_Click(object sender, EventArgs e)` — 按钮点击处理。

- [ ] **Step 1: 创建弹窗类**

创建 `src/ESQLNew/MainForm.ColumnMapDialog.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ESQLNew.Import;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class ColumnMapDialog : Form
        {
            private readonly DataGridView _grid;
            private readonly List<string> _fieldNames;

            public Dictionary<string, string> Result { get; private set; }

            public ColumnMapDialog(string table, IList<string> excelHeaders,
                IList<ColumnInfo> tableColumns, Dictionary<string, string> currentMap)
            {
                Text = "编辑列映射 - " + table;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(480, 420);

                _fieldNames = new List<string>();
                foreach (var c in tableColumns)
                    _fieldNames.Add(c.Name);

                _grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    RowHeadersVisible = false,
                    BackgroundColor = SystemColors.Window,
                    EditMode = DataGridViewEditMode.EditOnEnter
                };

                _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Excel 表头", ReadOnly = true });
                var fieldCol = new DataGridViewComboBoxColumn { HeaderText = "英文字段", DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton, FlatStyle = FlatStyle.Flat };
                foreach (var f in _fieldNames)
                    fieldCol.Items.Add(f);
                _grid.Columns.Add(fieldCol);

                for (int i = 0; i < excelHeaders.Count; i++)
                {
                    string header = excelHeaders[i] == null ? "" : excelHeaders[i].Trim();
                    if (header.Length == 0) continue;
                    string preset = null;
                    if (currentMap != null)
                        currentMap.TryGetValue(header, out preset);
                    _grid.Rows.Add(header, preset);
                }

                var save = new Button { Text = "保存", Width = 90, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Width = 90, DialogResult = DialogResult.Cancel };
                var flow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.RightToLeft,
                    WrapContents = false,
                    Height = 40
                };
                flow.Controls.Add(save);
                flow.Controls.Add(cancel);
                flow.BringToFront();

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
                layout.Controls.Add(_grid, 0, 0);
                layout.Controls.Add(flow, 0, 1);
                Controls.Add(layout);

                AcceptButton = save;
                CancelButton = cancel;
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (DialogResult == DialogResult.OK)
                    Result = CollectResult();
                base.OnFormClosing(e);
            }

            private Dictionary<string, string> CollectResult()
            {
                var headers = new List<string>();
                var fields = new List<string>();
                for (int i = 0; i < _grid.Rows.Count; i++)
                {
                    var header = _grid.Rows[i].Cells[0].Value as string;
                    var field = _grid.Rows[i].Cells[1].Value as string;
                    headers.Add(header);
                    fields.Add(field);
                }
                return ColumnMapStore.BuildMap(headers, fields);
            }
        }
    }
}
```

注意:文件顶部需加 `using ESQLNew.Core;`(CollectResult 用 ColumnMapStore)。

- [ ] **Step 2: 导入 Tab 加「编辑映射」按钮**

在 `src/ESQLNew/MainForm.Tabs.cs` 的 `BuildImportTab()` 中,`_previewButton` 创建(约第 176 行)附近新增字段与按钮。先在字段声明区(约第 43 行)加:

```csharp
        private Button _editMapButton;
```

在 `BuildImportTab()` 内 `_previewButton` 定义之后加:

```csharp
            _editMapButton = new Button { Text = "编辑映射", Dock = DockStyle.Fill };
```

在 `tableFlow` 构建处,将 `_tableRefreshButton` 加入布局的那一行改为加入 `_editMapButton`。当前 `tableFlow` 是 3 列(`_tableComboBox` / `_previewButton` / `_tableRefreshButton`)。改为 4 列,并在末尾加 `_editMapButton`:

```csharp
            var tableFlow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, Margin = new Padding(0) };
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.Controls.Add(_tableComboBox, 0, 0);
            tableFlow.Controls.Add(_previewButton, 1, 0);
            tableFlow.Controls.Add(_tableRefreshButton, 2, 0);
            tableFlow.Controls.Add(_editMapButton, 3, 0);
```

在 `BuildImportTab()` 的事件绑定处(`_chooseFileButton.Click += ...` 之后)加:

```csharp
            _editMapButton.Click += EditMapping_Click;
```

注意:实际布局可能已有 3 列 `tableFlow`(Task 之前实现),请按当前代码精确调整。

- [ ] **Step 3: 实现 EditMapping_Click**

在 `MainForm.Tabs.cs` 的 `Preview_Click` 方法之后追加:

```csharp
        private void EditMapping_Click(object sender, EventArgs e)
        {
            string path = _fileTextBox.Text.Trim();
            if (path.Length == 0)
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string table = _tableComboBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var headers = ExcelStreamReader.ReadHeaders(path);
                IList<ColumnInfo> cols = null;
                try
                {
                    cols = ImportEngine.GetTableColumns(CurrentConnectionString(), table);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "无法读取目标表字段:" + ex.Message + "\r\n仍可编辑,但英文字段下拉为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                var currentMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
                var dlg = new ColumnMapDialog(table, headers, cols ?? new List<ColumnInfo>(), currentMap);
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
                {
                    var maps = ColumnMapStore.Load(ColumnMapStore.ConfigPath);
                    maps[table] = dlg.Result;
                    ColumnMapStore.Save(ColumnMapStore.ConfigPath, maps);
                    _statusLabel.Text = "列映射已保存,共 " + dlg.Result.Count + " 列映射";
                    Preview_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "编辑映射失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
```

- [ ] **Step 4: 构建 + 测试**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`
Expected: 全部通过(51 + 3 = 54)。

- [ ] **Step 5: 提交**

```bash
git add src/ESQLNew/MainForm.ColumnMapDialog.cs src/ESQLNew/MainForm.Tabs.cs
git commit -m "feat: add column map editing dialog"
```

---

### Task 3: 收尾验证

**Files:**
- 无代码改动;验证与文档。

- [ ] **Step 1: 全量构建与测试**

Run: `dotnet build ESQLNew.sln` 与 `dotnet test ESQLNew.sln`
Expected: 0 错误,54 测试全绿。

- [ ] **Step 2: 静态检查**

Run: `rg "_editMapButton|ColumnMapDialog|BuildMap" src`
Expected: 新字段/类/方法被引用,无重复或残留。

- [ ] **Step 3: 代码审查**

按 superpowers:requesting-code-review 派发本功能全量 review(范围 = 本功能 2 个实现提交 + 文档提交的 diff)。

- [ ] **Step 4: 提交审查修复**

若审查发现问题,修复并提交 `fix: address map editor review findings`。

---

## Self-Review

**Spec coverage:**
- 编辑映射按钮:Task 2 Step 2 ✓
- 弹窗逐列(表头只读+英文字段下拉):Task 2 Step 1 ✓
- 候选从 GetTableColumns:Task 2 Step 1(构造参数接收 tableColumns)✓
- 未匹配默认空 / 已匹配预选:Task 2 Step 1(currentMap 预选)✓
- 保存只针对当前表:Task 2 Step 3(maps[table] = Result)✓
- 保存后重跑预览:Task 2 Step 3(Preview_Click 调用)✓
- 可测逻辑抽到 BuildMap:Task 1 ✓
- 未连接时下拉为空+提示:Task 2 Step 3 ✓
- 验收 6(测试全绿):Task 3 ✓

**Placeholder scan:** 无 TBD/TODO;所有步骤含完整代码。弹窗代码完整给出。

**Type consistency:** `ColumnMapDialog(table, excelHeaders, tableColumns, currentMap)` 构造在 Task 2 Step 1 定义、Step 3 使用一致;`ColumnMapStore.BuildMap(headers, fields)` 在 Task 1 定义、Task 2 Step 1 CollectResult 使用一致;`EditMapping_Click` 在 Step 2 绑定、Step 3 定义一致。