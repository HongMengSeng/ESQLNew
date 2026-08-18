# ESQLNew 匹配预览进度弹窗 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 匹配预览改为后台线程执行 + 独立进度弹窗,消除界面卡顿并显示步骤进度。

**Architecture:** 新增 `MainForm.ProgressDialog.cs`(非模态弹窗,进度条 0-4 + 步骤文字);`Preview_Click` 改为后台 `Task.Run` 执行四步匹配,每步经 `ShowStep`(BeginInvoke)更新弹窗,完成后 Invoke 回 UI 线程关弹窗并填充结果。

**Tech Stack:** .NET Framework 4.6.2 WinForms, xUnit 2.4.2。

## Global Constraints

- 目标框架 net462,LangVersion 7.3,代码不写任何注释,UI 中文。
- 构建:`dotnet build ESQLNew.sln` 0 错误;测试:`dotnet test ESQLNew.sln` 全绿(当前 54 个)。
- 匹配预览流程:读表头 → 读目标表结构(GetTableColumns)→ 列匹配(Map)→ 加载样例(ShowSample)。
- 进度弹窗用**非模态 `Show()`**;执行中禁用主窗口匹配预览按钮防重复点击;完成后恢复。
- 进度条 Minimum=0, Maximum=4;步骤文字:读取 Excel 表头 / 读取目标表结构 / 匹配列映射 / 加载样例数据。
- 后台线程不直接操作主窗口控件;进度与结果都经 BeginInvoke/Invoke 回 UI 线程。
- 分支:`main`,不 push(除非用户要求)。

---

### Task 1: 进度弹窗类

**Files:**
- Create: `src/ESQLNew/MainForm.ProgressDialog.cs`

**Interfaces:**
- Produces:
  - `partial class MainForm` 私有嵌套类 `ProgressDialog : Form`
    - 构造:`ProgressDialog(string title)`
    - 属性:`ProgressBar Bar`、`Label StepLabel`
    - 方法:`void ShowStep(int step, string text)` — 更新进度条值与步骤文字;此方法在 UI 线程调用(调用方保证已 BeginInvoke)

- [ ] **Step 1: 创建进度弹窗类**

创建 `src/ESQLNew/MainForm.ProgressDialog.cs`:

```csharp
using System.Drawing;
using System.Windows.Forms;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class ProgressDialog : Form
        {
            public ProgressBar Bar { get; private set; }
            public Label StepLabel { get; private set; }

            public ProgressDialog(string title)
            {
                Text = title;
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                ControlBox = false;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(360, 100);

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(10) };
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

                Bar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 4, Value = 0 };
                StepLabel = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Text = "正在准备…" };

                layout.Controls.Add(Bar, 0, 0);
                layout.Controls.Add(StepLabel, 0, 1);
                Controls.Add(layout);

                ShowInTaskbar = false;
            }

            public void ShowStep(int step, string text)
            {
                if (step < Bar.Minimum) step = Bar.Minimum;
                if (step > Bar.Maximum) step = Bar.Maximum;
                Bar.Value = step;
                StepLabel.Text = text;
            }
        }
    }
}
```

注意:`FixedToolWindow` + `ControlBox=false` 提供无关闭按钮的小窗体。此文件只创建类,不改其它文件。

- [ ] **Step 2: 构建验证**

Run: `dotnet build ESQLNew.sln`
Expected: 0 错误(类编译通过,虽暂未使用会有 CS0067 或 CS0414 警告——若出现未使用警告,忽略即可;`Bar`/`StepLabel` 为 public 属性不触发未使用警告)。

- [ ] **Step 3: 提交**

```bash
git add src/ESQLNew/MainForm.ProgressDialog.cs
git commit -m "feat: add progress dialog for preview"
```

---

### Task 2: Preview_Click 后台执行改造

**Files:**
- Modify: `src/ESQLNew/MainForm.Tabs.cs`

**Interfaces:**
- Consumes: `ProgressDialog`(Task 1)、`ExcelStreamReader.ReadHeaders`、`ImportEngine.GetTableColumns`、`ColumnMapStore.GetMap`、`ColumnMapper.Map`、`ShowMapping`、`ShowSample`、`IsValidTableName`、`TableExists`、`CurrentConnectionString`。
- Produces: 无新公开 API;`Preview_Click` 行为改为后台执行 + 进度弹窗。

- [ ] **Step 1: 替换 Preview_Click**

将 `src/ESQLNew/MainForm.Tabs.cs` 中 `Preview_Click` 整个方法体(从方法签名到结束 `}`)替换为:

```csharp
        private void Preview_Click(object sender, EventArgs e)
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
            if (!TableExists(table))
            {
                MessageBox.Show(this, "表名不存在或已更改,请重新选择", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _previewButton.Enabled = false;
            var dlg = new ProgressDialog("匹配预览中");
            dlg.Show(this);

            string connStr = CurrentConnectionString();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(1, "读取 Excel 表头")));
                    var headers = ExcelStreamReader.ReadHeaders(path);

                    dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(2, "读取目标表结构")));
                    var cols = ImportEngine.GetTableColumns(connStr, table);

                    dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(3, "匹配列映射")));
                    var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
                    var mappings = ColumnMapper.Map(headers, cols, fieldMap);

                    dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(4, "加载样例数据")));
                    BeginInvoke((System.Action)(() =>
                    {
                        ShowMapping(mappings);
                        ShowSample(mappings);
                        int matched = 0;
                        foreach (var m in mappings)
                            if (m.Matched) matched++;
                        _statusLabel.Text = string.Format("共 {0} 列,匹配 {1} 列,未匹配 {2} 列", mappings.Count, matched, mappings.Count - matched);
                        if (ColumnMapStore.LastLoadCorrupt)
                            _statusLabel.Text += "  列映射配置读取失败,已使用内置映射";
                        dlg.Close();
                        _previewButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    BeginInvoke((System.Action)(() =>
                    {
                        dlg.Close();
                        _previewButton.Enabled = true;
                        MessageBox.Show(this, ex.Message, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }
```

注意:需确保文件顶部有 `using ESQLNew.Core;`(已有)和 `using System.Linq;`(若 ShowSample/匹配计数用到——ShowSample 已存在,方法体未引用 LINQ;若编译报缺 using 则补)。`ShowMapping`/`ShowSample` 保持原样不动。

- [ ] **Step 2: 构建 + 测试**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`
Expected: 全部通过(54)。

- [ ] **Step 3: 提交**

```bash
git add src/ESQLNew/MainForm.Tabs.cs
git commit -m "feat: run preview matching in background with progress dialog"
```

---

### Task 3: 收尾验证

**Files:**
- 无代码改动;验证与文档。

- [ ] **Step 1: 全量构建与测试**

Run: `dotnet build ESQLNew.sln` 与 `dotnet test ESQLNew.sln`
Expected: 0 错误,54 测试全绿。

- [ ] **Step 2: 静态检查**

Run: `grep`(或 Select-String)`"ProgressDialog|ShowStep|_previewButton.Enabled"` in `src`
Expected: 弹窗类被引用,`_previewButton` 在后台执行时禁用/恢复。

- [ ] **Step 3: 代码审查**

按 superpowers:requesting-code-review 派发本功能全量 review(范围 = 本功能 2 个实现提交 + 文档提交的 diff)。

- [ ] **Step 4: 提交审查修复**

若审查发现问题,修复并提交 `fix: address preview progress review findings`。

---

## Self-Review

**Spec coverage:**
- 独立进度弹窗:Task 1 ✓
- 后台线程执行:Task 2(Task.Run)✓
- 4 步进度文字:Task 2(ShowStep 1-4)✓
- 完成后自动关闭+填充结果:Task 2(完成分支 dlg.Close + ShowMapping/ShowSample)✓
- 执行中禁用按钮防重复点击:Task 2(_previewButton.Enabled=false)✓
- 失败关闭弹窗+报错:Task 2(catch 分支)✓
- 非模态 Show + BeginInvoke 编组:Task 2 ✓
- 验收 6(测试全绿):Task 3 ✓

**Placeholder scan:** 无 TBD/TODO;所有步骤含完整代码。进度文字精确给出。

**Type consistency:** `ProgressDialog(title)`,`ShowStep(int,string)`,`Bar`,`StepLabel` 在 Task 1 定义、Task 2 使用一致。`Preview_Click` 完整替换,`ShowMapping/ShowSample` 保持原签名。后台线程用 `BeginInvoke` 更新弹窗与主窗口,与既有 Import 后台模式一致。