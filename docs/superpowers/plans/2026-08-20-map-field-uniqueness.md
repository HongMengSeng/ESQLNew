# 编辑映射英文字段唯一性限制实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 编辑列映射时,一个英文字段全局唯一——某行选中后其它行不可再选;打开时重复预置只留第一个。

**Architecture:** 在 `ColumnMapStore`(Core)新增两个纯静态方法(`ResolvePresets`、`AvailableFields`)便于单测;改造 `ColumnMapDialog`,把共享列级 ComboBoxColumn 改为每行独立 ComboBoxCell,维护 `_usedFields`,在字段变更时刷新其它行可用项。

**Tech Stack:** .NET Framework 4.6.2 WinForms, xUnit。

## Global Constraints

- net462, LangVersion 7.3
- 代码不写任何注释
- UI 文案为中文
- 现有 63 个测试保持通过
- 每个 Task 结束跑 `dotnet test ESQLNew.sln`,预期全绿
- 运行中 ESQLNew.exe 会锁 exe 导致 MSB3027/MSB3021 → 先 `Stop-Process -Name ESQLNew -Force`
- 仅修改本计划列出的文件

---

### Task 1: ColumnMapStore 纯逻辑方法 + 单元测试

**Files:**
- Modify: `src/ESQLNew/Core/ColumnMapStore.cs`
- Modify: `tests/ESQLNew.Tests/ColumnMapStoreTests.cs`

**Interfaces:**
- Produces (add to `ColumnMapStore`):
  - `public static List<KeyValuePair<string,string>> ResolvePresets(IList<string> excelHeaders, Dictionary<string,string> currentMap, IList<string> fieldNames)` — 返回 (表头, 预设字段) 对,重复字段只保留第一个出现,后续重复行预设为 null;空表头跳过。
  - `public static List<string> AvailableFields(IList<string> fieldNames, ISet<string> usedFields, string currentValue)` — 返回可用字段列表 = fieldNames − usedFields + currentValue(currentValue 非空且存在时加入)。
- Consumes: 无(纯静态,不依赖其它)

#### 逻辑说明

**ResolvePresets**:
- 遍历 excelHeaders;`header = h?.Trim()`;若空跳过。
- 从 currentMap 取 preset;若 preset 非空且不在 fieldNames → 置 null。
- 若 preset 非空且已在 `seen` 集合 → 该行 preset 置 null(重复,只留第一个)。
- 否则(首次)preset 保留,加入 `seen`。
- 返回 List<(header, preset)>。

**AvailableFields**:
- 结果 = fieldNames 中所有不在 usedFields 的项;若 currentValue 非空且不在结果中(即 currentValue 在 usedFields 里)则追加 currentValue。
- 返回按 fieldNames 顺序的列表。

- [ ] **Step 1: 写失败测试**

在 `tests/ESQLNew.Tests/ColumnMapStoreTests.cs` 的 `ColumnMapStoreTests` 类末尾(现有 3 个 BuildMap 测试后)新增:

```csharp
[Fact]
public void ResolvePresets_KeepsFirstDuplicateAndNullsRest()
{
    var headers = new List<string> { "分类", "其它", "分类" };
    var currentMap = new Dictionary<string, string>
    {
        { "分类", "category" },
        { "其它", "other" }
    };
    var fieldNames = new List<string> { "category", "other" };
    var result = ColumnMapStore.ResolvePresets(headers, currentMap, fieldNames);
    Assert.Equal(3, result.Count);
    Assert.Equal("category", result[0].Value);
    Assert.Equal("other", result[1].Value);
    Assert.Null(result[2].Value);
}

[Fact]
public void ResolvePresets_SkipsBlankHeaderAndStaleField()
{
    var headers = new List<string> { "  ", "旧字段列" };
    var currentMap = new Dictionary<string, string> { { "旧字段列", "not_exist_field" } };
    var fieldNames = new List<string> { "category" };
    var result = ColumnMapStore.ResolvePresets(headers, currentMap, fieldNames);
    Assert.Equal(1, result.Count);
    Assert.Equal("旧字段列", result[0].Key);
    Assert.Null(result[0].Value);
}

[Fact]
public void AvailableFields_ExcludesUsed_KeepsCurrentValue()
{
    var fieldNames = new List<string> { "a", "b", "c" };
    var used = new HashSet<string> { "b" };
    var result = ColumnMapStore.AvailableFields(fieldNames, used, "b");
    Assert.Equal(3, result.Count);
    Assert.Contains("a", result);
    Assert.Contains("b", result);
    Assert.Contains("c", result);
    Assert.DoesNotContain(used.Except(new[] { "b" }), x => !result.Contains(x));
}

[Fact]
public void AvailableFields_DropsUsedOthers_NoCurrent()
{
    var fieldNames = new List<string> { "a", "b", "c" };
    var used = new HashSet<string> { "b", "c" };
    var result = ColumnMapStore.AvailableFields(fieldNames, used, null);
    Assert.Equal(1, result.Count);
    Assert.Equal("a", result[0]);
}
```

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ResolvePresets|FullyQualifiedName~AvailableFields"`
Expected: 编译失败(方法不存在)→ 红。

- [ ] **Step 3: 实现**

在 `src/ESQLNew/Core/ColumnMapStore.cs` 的 `BuildMap` 方法后新增:

```csharp
public static List<KeyValuePair<string, string>> ResolvePresets(IList<string> excelHeaders,
    Dictionary<string, string> currentMap, IList<string> fieldNames)
{
    var result = new List<KeyValuePair<string, string>>();
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var h in excelHeaders)
    {
        string header = h == null ? "" : h.Trim();
        if (header.Length == 0) continue;
        string preset = null;
        if (currentMap != null)
            currentMap.TryGetValue(header, out preset);
        if (preset != null && !fieldNames.Contains(preset))
            preset = null;
        if (preset != null && !seen.Add(preset))
            preset = null;
        result.Add(new KeyValuePair<string, string>(header, preset));
    }
    return result;
}

public static List<string> AvailableFields(IList<string> fieldNames, ISet<string> usedFields, string currentValue)
{
    var result = new List<string>();
    foreach (var f in fieldNames)
        if (!usedFields.Contains(f))
            result.Add(f);
    if (!string.IsNullOrEmpty(currentValue) && !result.Contains(currentValue))
        result.Add(currentValue);
    return result;
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ResolvePresets|FullyQualifiedName~AvailableFields"`
Expected: PASS。

- [ ] **Step 5: 运行全部测试 + 提交**

Run: `dotnet test ESQLNew.sln` → 全绿(原 63 + 新增 4 = 67)。
Commit:
```bash
git add src/ESQLNew/Core/ColumnMapStore.cs tests/ESQLNew.Tests/ColumnMapStoreTests.cs
git commit -m "feat: add preset resolution and available-fields helpers"
```

---

### Task 2: 改造 ColumnMapDialog 使用每行唯一字段

**Files:**
- Modify: `src/ESQLNew/MainForm.ColumnMapDialog.cs`

**Interfaces:**
- Consumes: `ColumnMapStore.ResolvePresets(excelHeaders, currentMap, _fieldNames)`、`ColumnMapStore.AvailableFields(_fieldNames, _usedFields, currentValue)`(Task 1)
- Produces: 无新公开接口(仍用 `Result` 属性)

#### 改造点

1. **字段声明**:新增 `private readonly HashSet<string> _usedFields;`。

2. **构造函数**:
   - 用 `ColumnMapStore.ResolvePresets(excelHeaders, currentMap, _fieldNames)` 替代原第 57-67 行的行构建循环,得到 (header, preset) 列表。
   - `_usedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)`;遍历 ResolvePresets 结果,preset 非空则加入 `_usedFields`。
   - 改为**每行独立 ComboBoxCell**:
     - 仍用 `DataGridViewComboBoxColumn` 作为列(便于编辑控件),但不再在列级 `fieldCol.Items` 添加所有字段。而是在每行创建 `DataGridViewComboBoxCell`,设置 `cell.Items` 为该行可用字段 `AvailableFields(_fieldNames, _usedFields, preset)`,`cell.Value = preset`,放入行的第二格。
   - 保留 `EditingControlShowing` 的 `DropDownStyle = DropDownList`。
   - 新增 `_grid.CellValueChanged += (s,e) => { if (e.ColumnIndex == 1) RefreshUsedFields(); };` 和 `_grid.CurrentCellDirtyStateChanged += (s,e) => { if (_grid.IsCurrentCellDirty && _grid.CurrentCell.ColumnIndex == 1) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };`(确保编辑立即提交,触发 CellValueChanged)。
   - 新增方法 `RefreshUsedFields()`:重算 `_usedFields`(遍历每行 Cells[1].Value 非空则加入,用 OrdinalIgnoreCase),然后为每行更新 ComboBoxCell 的 Items(用 `AvailableFields(_fieldNames, _usedFields, 该行当前值)`)。注意:更新 Items 会重置 Value,需先保存该行当前值,设置 Items 后再恢复 Value。

3. **Remove Preset-CellValueChanged 时避免递归**:`RefreshUsedFields` 更新 Items 不应再次触发 CellValueChanged 造成死循环——设置 cell.Value 时若值未变化则不触发,或加标志位 `_updating` 保护。实现者需确保刷新逻辑幂等且无递归。

4. **`CollectResult()` 不变**。

#### 关键实现模板(供参考)

```csharp
private void RefreshUsedFields()
{
    _updating = true;
    try
    {
        _usedFields.Clear();
        for (int i = 0; i < _grid.Rows.Count; i++)
        {
            var val = _grid.Rows[i].Cells[1].Value as string;
            if (!string.IsNullOrEmpty(val))
                _usedFields.Add(val);
        }
        for (int i = 0; i < _grid.Rows.Count; i++)
        {
            var cell = _grid.Rows[i].Cells[1] as DataGridViewComboBoxCell;
            if (cell == null) continue;
            string current = _grid.Rows[i].Cells[1].Value as string;
            var items = ColumnMapStore.AvailableFields(_fieldNames, _usedFields, current);
            cell.Items.Clear();
            foreach (var it in items) cell.Items.Add(it);
            if (cell.Value == null || !items.Contains(cell.Value as string))
                cell.Value = current;
        }
    }
    finally
    {
        _updating = false;
    }
}
```

实现者需按上述思路补全(含 `_updating` 标志、行构建改为 ResolvePresets + 每行 ComboBoxCell)。

#### 测试

UI 交互(net462 WinForms DataGridView)不便单测;逻辑已由 Task 1 的纯方法测试覆盖。验证方式:构建 0 错误 + 现有测试全绿 + 手动逻辑验证(用 probe 或直接启动)。

- [ ] **Step 1: 实现上述改造**

- [ ] **Step 2: 构建 + 测试**

Run: `dotnet build ESQLNew.sln -c Release`(0 错误)→ `dotnet test ESQLNew.sln`(全绿,67 个)

- [ ] **Step 3: 手动逻辑验证**

用临时 probe(参考 C:\Users\Kuade\AppData\Local\Temp\opencode\probe,probe.csproj 需 Compile-Include 相关源文件)调用 `ColumnMapStore.ResolvePresets` 和 `AvailableFields`,打印结果验证:
- 重复字段只留第一个
- AvailableFields 排除已用、保留当前值
- 或用 probe 直接实例化 ColumnMapDialog(若可)验证打开时重复预置清空

- [ ] **Step 4: 提交**

```bash
git add src/ESQLNew/MainForm.ColumnMapDialog.cs
git commit -m "feat: enforce unique target fields in map editor"
```
