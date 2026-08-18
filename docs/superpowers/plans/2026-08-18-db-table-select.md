# ESQLNew 数据库/表选择功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 连接成功后从下拉列表选库、并自动拉取该库所有表供导入选择。

**Architecture:** 新增 `Core/DbMetadata` 静态类封装 MySQL 元数据查询(SHOW DATABASES / information_schema.TABLES,含系统库过滤);`MySqlConnectionBuilder.Build` 支持空库名连接串;连接 Tab 与导入 Tab 的文本框改为可搜索 ComboBox,测试连接成功后自动填充库列表,选库自动填充表列表,并提供手动刷新按钮与"显示系统库"开关。

**Tech Stack:** .NET Framework 4.6.2 WinForms, MySql.Data 8.0.33, xUnit 2.4.2。

## Global Constraints

- 目标框架 net462,LangVersion 7.3,代码不写任何注释,UI 中文。
- 构建:`dotnet build ESQLNew.sln` 0 错误;测试:`dotnet test ESQLNew.sln` 全绿(当前 24 个测试)。
- 每个任务一个独立提交,提交信息见各任务。
- 库列表查询用 `SHOW DATABASES`;表列表查询用 `information_schema.TABLES` 且仅 BASE TABLE、参数化、ORDER BY TABLE_NAME。
- 系统库集合:`mysql`、`information_schema`、`performance_schema`、`sys`,比较大小写不敏感。
- 库/表名输入仍沿用现有 `^[A-Za-z0-9_]+$` 校验;不改变现有连接/预览/导入/日志行为。
- 仓库位于 `C:\Users\Kuade\Desktop\Exit`,分支 `feature/excel-import`,不 push。

---

### Task 1: DbMetadata 元数据查询 + 空库连接串 + 配置开关

**Files:**
- Create: `src/ESQLNew/Core/DbMetadata.cs`
- Modify: `src/ESQLNew/Core/MySqlConnectionBuilder.cs`
- Modify: `src/ESQLNew/Core/AppConfig.cs`
- Test: `tests/ESQLNew.Tests/DbMetadataTests.cs`, `tests/ESQLNew.Tests/MySqlConnectionBuilderTests.cs`, `tests/ESQLNew.Tests/AppConfigTests.cs`

**Interfaces:**
- Produces:
  - `static class ESQLNew.Core.DbMetadata`
    - `static IList<string> GetDatabases(string connStr)` — 执行 `SHOW DATABASES`,返回所有库名(未过滤系统库)。
    - `static IList<string> GetTables(string connStr, string database)` — 参数化查询 `SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME`,返回表名列表。
    - `static bool IsSystemDatabase(string name)` — 返回 name 是否系统库(mysql/information_schema/performance_schema/sys,忽略大小写与两端空白)。
    - GetDatabases/GetTables 的底层异常包装为 `InvalidOperationException`,消息为中文("无法读取数据库列表:..." / "无法读取表列表:...")。
  - `MySqlConnectionBuilder.Build(server, port, user, password, database)` — database 为空或空白时连接串不含 `Database=...` 段。
  - `AppConfig.ShowSystemDatabases` — `bool` 属性,默认 `false`。

- [ ] **Step 1: 写失败的测试**

`tests/ESQLNew.Tests/DbMetadataTests.cs` 新建:

```csharp
using System;
using System.Collections.Generic;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class DbMetadataTests
    {
        [Theory]
        [InlineData("mysql")]
        [InlineData("information_schema")]
        [InlineData("performance_schema")]
        [InlineData("sys")]
        [InlineData("MYSQL")]
        [InlineData("  sys  ")]
        public void IsSystemDatabase_ReturnsTrue(string name)
        {
            Assert.True(DbMetadata.IsSystemDatabase(name));
        }

        [Theory]
        [InlineData("testdb")]
        [InlineData("esqlnew")]
        [InlineData("sales")]
        [InlineData("")]
        public void IsSystemDatabase_ReturnsFalse(string name)
        {
            Assert.False(DbMetadata.IsSystemDatabase(name));
        }
    }
}
```

`tests/ESQLNew.Tests/MySqlConnectionBuilderTests.cs` 追加:

```csharp
        [Fact]
        public void Build_EmptyDatabase_OmitsDatabasePart()
        {
            var cs = ESQLNew.Core.MySqlConnectionBuilder.Build(
                "192.168.201.112", 3306, "root", "pwd", "");
            Assert.DoesNotContain("Database=", cs);
        }

        [Fact]
        public void Build_WhitespaceDatabase_OmitsDatabasePart()
        {
            var cs = ESQLNew.Core.MySqlConnectionBuilder.Build(
                "192.168.201.112", 3306, "root", "pwd", "   ");
            Assert.DoesNotContain("Database=", cs);
        }
```

`tests/ESQLNew.Tests/AppConfigTests.cs` 追加:

```csharp
        [Fact]
        public void Default_ShowSystemDatabases_False()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_sysdb_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = AppConfig.Load(path);
            Assert.False(cfg.ShowSystemDatabases);
        }
```

- [ ] **Step 2: 跑测试确认失败**

Run: `dotnet test ESQLNew.sln`
Expected: FAIL(DbMetadata 不存在、Build 仍含 Database=、ShowSystemDatabases 未定义)。

- [ ] **Step 3: 实现 DbMetadata**

创建 `src/ESQLNew/Core/DbMetadata.cs`:

```csharp
using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;

namespace ESQLNew.Core
{
    public static class DbMetadata
    {
        private static readonly string[] SystemDatabases = { "mysql", "information_schema", "performance_schema", "sys" };

        public static bool IsSystemDatabase(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            var trimmed = name.Trim();
            foreach (var s in SystemDatabases)
                if (string.Equals(trimmed, s, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        public static IList<string> GetDatabases(string connStr)
        {
            var result = new List<string>();
            try
            {
                using (var conn = new MySqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand("SHOW DATABASES", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            result.Add(reader.GetString(0));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取数据库列表:" + ex.Message, ex);
            }
            return result;
        }

        public static IList<string> GetTables(string connStr, string database)
        {
            var result = new List<string>();
            try
            {
                using (var conn = new MySqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand(
                        "SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME",
                        conn))
                    {
                        cmd.Parameters.AddWithValue("@db", database);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                result.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取表列表:" + ex.Message, ex);
            }
            return result;
        }
    }
}
```

- [ ] **Step 4: 修改 MySqlConnectionBuilder 支持空库名**

将 `src/ESQLNew/Core/MySqlConnectionBuilder.cs` 的 Build 改为:

```csharp
namespace ESQLNew.Core
{
    public static class MySqlConnectionBuilder
    {
        public static string Build(string server, int port, string user, string password, string database)
        {
            string dbPart = string.IsNullOrWhiteSpace(database) ? "" : ";Database=" + database;
            return string.Format(
                "Server={0};Port={1};Uid={2};Pwd={3}{4};SslMode=None;AllowLoadLocalInfile=false;CharSet=utf8mb4;",
                server, port, user, password, dbPart);
        }
    }
}
```

注意:原测试 `Build_IncludesAllParts` 断言 `Database=testdb` 仍须通过(非空库名时 dbPart 为 `;Database=testdb`)。

- [ ] **Step 5: 修改 AppConfig 加开关**

在 `src/ESQLNew/Core/AppConfig.cs` 的 `LogRetentionDays` 之后追加:

```csharp
        public bool ShowSystemDatabases { get; set; } = false;
```

- [ ] **Step 6: 跑测试确认通过**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`
Expected: 全部通过(24 + 新增约 9 = 约 33 个)。

- [ ] **Step 7: 提交**

```bash
git add src/ESQLNew/Core/DbMetadata.cs src/ESQLNew/Core/MySqlConnectionBuilder.cs src/ESQLNew/Core/AppConfig.cs tests/ESQLNew.Tests/DbMetadataTests.cs tests/ESQLNew.Tests/MySqlConnectionBuilderTests.cs tests/ESQLNew.Tests/AppConfigTests.cs
git commit -m "feat: db metadata queries + no-db connection + system db toggle config"
```

---

### Task 2: 连接 Tab 库下拉(自动拉取 + 刷新 + 显示系统库开关)

**Files:**
- Modify: `src/ESQLNew/MainForm.Tabs.cs`(连接 Tab 构建、LoadConfig、UpdateModeFields、CurrentConnectionString、TestConnection、SaveConfig)

**Interfaces:**
- Consumes: `DbMetadata.GetDatabases(connStr)`, `DbMetadata.IsSystemDatabase(name)`, `AppConfig.ShowSystemDatabases`, `CurrentConnectionString()`(改为 `ComboBox` 取值)。
- Produces:
  - 字段 `_databaseComboBox`(ComboBox,`DropDown` 可输入可下拉、AutoComplete)替代 `_databaseTextBox`。
  - 字段 `_dbRefreshButton`(刷新库列表)、`_showSysDbCheckBox`(显示系统库)。
  - 方法 `void ReloadDatabases()` — 拉取并填充库下拉(内部处理连接失败提示)。
  - 方法 `void ReloadTables(string database)` — 拉取并填充导入 Tab 表下拉(内部处理失败提示)。
  - 事件:测试连接成功后调用 `ReloadDatabases()`;`_databaseComboBox.SelectedIndexChanged` 触发 `ReloadTables(选中库)`;`_dbRefreshButton.Click` 调用 `ReloadDatabases()`;`_showSysDbCheckBox.CheckedChanged` 保存配置并调用 `ReloadDatabases()`。

- [ ] **Step 1: 替换字段声明**

在 `src/ESQLNew/MainForm.Tabs.cs` 第 23 行附近,将:

```csharp
        private TextBox _databaseTextBox;
```

替换为:

```csharp
        private ComboBox _databaseComboBox;
        private Button _dbRefreshButton;
        private CheckBox _showSysDbCheckBox;
        private ComboBox _tableComboBox;
        private Button _tableRefreshButton;
```

注意:`_tableComboBox`/`_tableRefreshButton` 字段在此声明、Task 3 才在 BuildImportTab 中实例化;本任务 `ReloadTables` 方法已引用它们,字段必须先存在才能编译。`_tableTextBox` 字段保留至 Task 3 移除(本任务不触碰导入 Tab 构建)。

- [ ] **Step 2: 构建连接 Tab 控件**

在 `BuildConnectionGroup()` 中,将第 328 行:

```csharp
            _databaseTextBox = new TextBox { Dock = DockStyle.Fill };
```

替换为:

```csharp
            _databaseComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend
            };
            _dbRefreshButton = new Button { Text = "刷新库", Width = 64 };
            _showSysDbCheckBox = new CheckBox
            {
                Text = "显示系统库",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 0)
            };
            var dbFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            dbFlow.Controls.Add(_databaseComboBox);
            dbFlow.Controls.Add(_dbRefreshButton);
            dbFlow.Controls.Add(_showSysDbCheckBox);
```

将第 355 行:

```csharp
            layout.Controls.Add(_databaseTextBox, 1, 4);
```

替换为:

```csharp
            layout.Controls.Add(dbFlow, 1, 4);
```

- [ ] **Step 3: 更新 UpdateModeFields**

将第 435 行:

```csharp
            _databaseTextBox.Enabled = form;
```

替换为:

```csharp
            _databaseComboBox.Enabled = form;
            _dbRefreshButton.Enabled = form;
            _showSysDbCheckBox.Enabled = form;
```

- [ ] **Step 4: 更新 CurrentConnectionString / SaveConfig / LoadConfig**

将第 449 行:

```csharp
                _databaseTextBox.Text.Trim());
```

替换为:

```csharp
                _databaseComboBox.Text.Trim());
```

将第 498 行:

```csharp
                Database = _databaseTextBox.Text.Trim(),
```

替换为:

```csharp
                Database = _databaseComboBox.Text.Trim(),
```

将第 416 行:

```csharp
            _databaseTextBox.Text = cfg.Database ?? "";
```

替换为:

```csharp
            _databaseComboBox.Text = cfg.Database ?? "";
            _showSysDbCheckBox.Checked = cfg.ShowSystemDatabases;
```

**SaveConfig 必改点:** `SaveConfig()` 中构造新 `AppConfig` 的初始化器必须新增一行 `ShowSystemDatabases = _showSysDbCheckBox.Checked,`(紧邻 `KeepLogs` 之后),否则点"保存配置"会把开关状态回写为默认 `false`。

- [ ] **Step 5: 在 TestConnection 成功后自动拉取库列表**

在 `TestConnection()` 的 `BeginInvoke` 回调中,`error == null` 分支内(第 475 行附近,`MessageBox.Show` 之后)追加:

```csharp
                    ReloadDatabases();
```

- [ ] **Step 6: 添加 ReloadDatabases / ReloadTables 方法**

在 `UpdateModeFields()` 方法之后追加:

```csharp
        private void ReloadDatabases()
        {
            try
            {
                var connStr = CurrentConnectionString();
                var all = DbMetadata.GetDatabases(connStr);
                var list = new List<string>();
                foreach (var db in all)
                    if (_showSysDbCheckBox.Checked || !DbMetadata.IsSystemDatabase(db))
                        list.Add(db);
                string current = _databaseComboBox.Text.Trim();
                _databaseComboBox.Items.Clear();
                foreach (var db in list)
                    _databaseComboBox.Items.Add(db);
                bool found = false;
                foreach (var db in list)
                    if (string.Equals(db, current, StringComparison.OrdinalIgnoreCase))
                        found = true;
                if (!found)
                {
                    if (list.Count > 0)
                        _databaseComboBox.Text = "";
                    if (_databaseComboBox.Items.Count > 0)
                        _databaseComboBox.SelectedIndex = -1;
                }
                else
                {
                    _databaseComboBox.Text = current;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "刷新库列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ReloadTables(string database)
        {
            if (string.IsNullOrWhiteSpace(database)) return;
            try
            {
                var tables = DbMetadata.GetTables(CurrentConnectionString(), database.Trim());
                _tableComboBox.Items.Clear();
                foreach (var t in tables)
                    _tableComboBox.Items.Add(t);
                _tableComboBox.Text = "";
            }
            catch (Exception ex)
            {
                _tableComboBox.Items.Clear();
                _tableComboBox.Text = "";
                MessageBox.Show(this, ex.Message, "刷新表列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
```

- [ ] **Step 7: 挂接事件(选库即拉表、刷新按钮、显示系统库开关)**

在 `BuildConnectionTab()` 的 `_saveButton.Click += ...` 之后追加:

```csharp
            _dbRefreshButton.Click += (s, e) => ReloadDatabases();
            _showSysDbCheckBox.CheckedChanged += (s, e) =>
            {
                if (_loadingConfig) return;
                try
                {
                    var cfg = AppConfig.Load(ConfigPath);
                    cfg.ShowSystemDatabases = _showSysDbCheckBox.Checked;
                    cfg.Save(ConfigPath);
                }
                catch
                {
                }
                ReloadDatabases();
            };
            _databaseComboBox.SelectedIndexChanged += (s, e) =>
            {
                if (_databaseComboBox.SelectedItem != null)
                    ReloadTables(_databaseComboBox.SelectedItem.ToString());
            };
```

注意:Task 1 已改过 `_keepLogsCheckBox` 的 CheckedChanged 用 `_loadingConfig` 防抖,本步骤保持一致模式。需确认 `_loadingConfig` 字段已存在(Task 1 未删除,仍在文件中)。

- [ ] **Step 8: 构建 + 测试 + 提交**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`(全绿)。
Commit:

```bash
git add src/ESQLNew/MainForm.Tabs.cs
git commit -m "feat: database dropdown with refresh and system-db toggle"
```

---

### Task 3: 导入 Tab 表下拉 + 校验

**Files:**
- Modify: `src/ESQLNew/MainForm.Tabs.cs`(导入 Tab 构建、ChooseFile_Click、Preview_Click、StartImport_Click、IsValidTableName 调用处)

**Interfaces:**
- Consumes: `ReloadTables(string)`(Task 2 产出)、`_databaseComboBox`(Task 2 产出)、`DbMetadata`。
- Produces:
  - 字段 `_tableComboBox`(ComboBox,`DropDown` 可输入可下拉)替代 `_tableTextBox`。
  - 字段 `_tableRefreshButton`(刷新表)。
  - 方法 `bool TableExists(string table)` — 当前表下拉非空时判断 table 是否在列表项中(大小写不敏感);下拉为空时返回 true(不做存在性断言)。

- [ ] **Step 1: 删除残留字段声明**

在 `src/ESQLNew/MainForm.Tabs.cs` 第 36 行附近,删除整行:

```csharp
        private TextBox _tableTextBox;
```

注意:字段 `_tableComboBox`/`_tableRefreshButton` 已在 Task 2 Step 1 声明,此处仅删除残留的 `_tableTextBox` 声明行。确保最终代码无重复 `_tableComboBox`/`_tableRefreshButton` 声明、无残留 `_tableTextBox` 引用。

- [ ] **Step 2: 构建导入 Tab 控件**

在 `BuildImportTab()` 中,将第 145 行:

```csharp
            _tableTextBox = new TextBox { Dock = DockStyle.Fill };
```

替换为:

```csharp
            _tableComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend
            };
            _tableRefreshButton = new Button { Text = "刷新表", Dock = DockStyle.Fill };
```

将第 150-154 行的 tableFlow 构建改为:

```csharp
            var tableFlow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Margin = new Padding(0) };
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.Controls.Add(_tableComboBox, 0, 0);
            tableFlow.Controls.Add(_previewButton, 1, 0);
            tableFlow.Controls.Add(_tableRefreshButton, 2, 0);
```

注意:`_previewButton` 已在 task 之前的代码中创建(第 144 行),此步仅调整 tableFlow 布局。原 tableFlow 第 154 行 `tableFlow.Controls.Add(_tableTextBox, 0, 0)` 替换为 `_tableComboBox`;`_tableRefreshButton` 必须加入布局(第三列),否则按钮不可见。

- [ ] **Step 3: 更新 ChooseFile_Click**

将第 532、538 行:

```csharp
                _tableTextBox.Text = "";
```
```csharp
                        _tableTextBox.Text = names[0];
```

替换为:

```csharp
                _tableComboBox.Text = "";
```
```csharp
                        _tableComboBox.Text = names[0];
```

- [ ] **Step 4: 更新 Preview_Click / StartImport_Click 取表名**

将第 555、587 行:

```csharp
            string table = _tableTextBox.Text.Trim();
```

替换为:

```csharp
            string table = _tableComboBox.Text.Trim();
```

- [ ] **Step 5: 更新启用/禁用逻辑**

将第 600、633 行:

```csharp
            _tableTextBox.Enabled = false;
```
```csharp
                _tableTextBox.Enabled = true;
```

替换为:

```csharp
            _tableComboBox.Enabled = false;
            _tableRefreshButton.Enabled = false;
```
```csharp
                _tableComboBox.Enabled = true;
                _tableRefreshButton.Enabled = true;
```

- [ ] **Step 6: 添加 TableExists 校验方法**

在 `IsValidTableName` 方法之后追加:

```csharp
        private bool TableExists(string table)
        {
            if (_tableComboBox.Items.Count == 0) return true;
            foreach (var item in _tableComboBox.Items)
                if (string.Equals(item.ToString(), table, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }
```

- [ ] **Step 7: 在预览/导入前加表存在性校验**

在 `Preview_Click` 中,`IsValidTableName` 校验之后(`if (!IsValidTableName(table))` 块之后)追加:

```csharp
            if (!TableExists(table))
            {
                MessageBox.Show(this, "表名不存在或已更改,请重新选择", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
```

在 `StartImport_Click` 中,同样的位置追加相同校验块。

- [ ] **Step 8: 挂接刷新表按钮**

在 `BuildImportTab()` 的 `_chooseFileButton.Click += ...` 之后追加:

```csharp
            _tableRefreshButton.Click += (s, e) =>
            {
                if (_databaseComboBox.SelectedItem != null)
                    ReloadTables(_databaseComboBox.SelectedItem.ToString());
                else if (!string.IsNullOrWhiteSpace(_databaseComboBox.Text))
                    ReloadTables(_databaseComboBox.Text.Trim());
                else
                    MessageBox.Show(this, "请先选择数据库", "刷新表列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
```

- [ ] **Step 9: 构建 + 测试 + 提交**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`(全绿)。
Commit:

```bash
git add src/ESQLNew/MainForm.Tabs.cs
git commit -m "feat: target table dropdown with existence validation"
```

---

### Task 4: 收尾验证

**Files:**
- 无代码改动;仅验证与检查。

**Interfaces:**
- 验证 Task 1-3 产出的完整链路。

- [ ] **Step 1: 全量构建与测试**

Run: `dotnet build ESQLNew.sln` 与 `dotnet test ESQLNew.sln`
Expected: 0 错误,全部测试通过。

- [ ] **Step 2: 静态检查字段一致性**

Run: `rg "_databaseTextBox|_tableTextBox" src` 与 `rg "_databaseComboBox|_tableComboBox" src`
Expected: 无旧字段残留;新字段引用存在。

- [ ] **Step 3: 代码审查请求**

按 superpowers:requesting-code-review 派发全量 review(范围 = 本功能 3 个提交的 diff)。
Expected: 无 Critical/Important 问题;或记录待办。

- [ ] **Step 4: 提交任何审查修复**

若审查发现问题,修复并提交 `fix: address db/table select review findings`。

---

## Self-Review

**Spec coverage:**
- 需求 1(选库选表):Task 2 库下拉 + Task 3 表下拉 ✓
- 需求 2(自动+手动刷新):Task 2 Step 5(测试后自动)+ Step 7(刷新按钮)+ Task 3 Step 8(刷新表)✓
- 需求 3(可搜索下拉):Task 2/3 ComboBox AutoComplete ✓
- 需求 4(有权限库+可搜索+隐藏系统库开关):Task 1 IsSystemDatabase + Task 2 _showSysDbCheckBox ✓
- 需求 5(选库即拉表):Task 2 Step 7 SelectedIndexChanged ✓
- 需求 6(无库连接):Task 1 MySqlConnectionBuilder 空库名 ✓
- 验收 5(表存在性校验):Task 3 Step 6/7 ✓
- 验收 6(不回归):Task 4 ✓
- 测试要求:Task 1 单元测试(IsSystemDatabase/Build/AppConfig)、Task 2/3 UI 手工验收 ✓

**Placeholder scan:** 无 TBD/TODO;所有步骤含完整代码与命令。ReloadDatabases/ReloadTables/TableExists 在 Task 2/3 内定义并被同任务或后续任务引用,无悬空引用。

**Type consistency:** `_databaseComboBox`/`_tableComboBox` 在 Task 2/3 声明与全部引用点一致;`ReloadTables(string)` 签名在 Task 2 定义、Task 3 复用一致;`DbMetadata.GetDatabases/GetTables/IsSystemDatabase` 签名在 Task 1 定义、Task 2/3 使用一致;`AppConfig.ShowSystemDatabases` 在 Task 1 定义、Task 2 使用一致。