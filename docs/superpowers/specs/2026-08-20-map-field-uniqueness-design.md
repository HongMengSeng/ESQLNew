# 编辑映射英文字段唯一性限制设计

日期:2026-08-20
状态:已批准

## 背景与问题

`ColumnMapDialog`(编辑列映射弹窗)当前允许同一英文字段被多行 Excel 表头重复选择,导致生成重复目标字段(如两个"分类"都映射到 `category`),进而引发 INSERT 重复列错误。需求:编辑映射时,一个英文字段**全局唯一**——某行选中后,其它行不可再选;打开时若已有重复预置,只保留第一个,后续重复行清空。

## 需求

1. 一个英文字段全局唯一:某行选中字段 X,其它行下拉中 X 移除(不可选)。
2. 打开弹窗时,若现有配置里多个表头映射到同一字段(旧数据),只保留第一个,后续重复行清空(可选)。
3. 未使用字段保持可选(现状不变)。
4. 用户可在下拉中把某行改回空(取消映射),释放字段。

## 架构

### 组件改造:`src/ESQLNew/MainForm.ColumnMapDialog.cs`

**核心:每行独立 `DataGridViewComboBoxCell`** — 当前用共享列级 `fieldCol.Items`(所有行同一选项集)。改为每行独立的 ComboBoxCell,以便动态控制每行可用项。

**维护状态**:`HashSet<string> _usedFields` 记录已占用字段。

**初始化逻辑**:
- 遍历 excelHeaders 构建行;已映射字段(preset)若尚未在 `_usedFields` 中 → 保留并加入 `_usedFields`;若已在 `_usedFields`(重复)→ 该行预设清空,不加入。
- 每行 ComboBoxCell.Items = `_fieldNames − _usedFields + 该行当前值`(该行保留自己已选字段,其它行已用字段被移除)。

**编辑交互**:
- 用 `DataGridViewComboBoxColumn` 的 `CellValueChanged`(或配合 `CurrentCellDirtyStateChanged` 在编辑结束提交时)检测字段变更。
- 当某行字段从 A 变为 B:
  - 旧值 A:若其它行没有使用 A → 从 `_usedFields` 移除。
  - 新值 B:加入 `_usedFields`(若 B 非空)。
  - 刷新所有其它行:每行 Items = `_fieldNames − _usedFields + 该行当前值`;若某行当前值已被其它行占用(理论不应发生,因刷新保持唯一),清空。
- 值改为空 → 从 `_usedFields` 释放。

### 关键公式

```
每行可选项 = _fieldNames − _usedFields + 该行当前选中字段
```

`_usedFields` 指其它行已占用的字段(不含本行当前值)。

### 交互细节

- 保留现有 `EditingControlShowing` 里的 `DropDownStyle = DropDownList`(只读下拉)。
- `CollectResult()` 不变(从每行 Cells[0]/Cells[1] 收集)。

## 数据流

```
打开弹窗 → 构建行(重复预置去重,只留第一个)
  → 初始化 _usedFields 与每行 Items
  → 用户编辑某行字段 → 更新 _usedFields → 刷新其它行 Items
  → 保存 → CollectResult 收集(唯一)
```

## 错误处理

- 用户尝试把某行设成其它行已用字段:不可选(已从 Items 移除),自然无法选择。
- 若字段被释放后其它行可重新选择。

## 测试

- UI 交互(net462 WinForms DataGridViewComboBox)不便单测,用构建验证 + 逻辑验证。
- 若可行,把"去重预置"与"计算每行可选项"提取为纯静态方法便于单测:
  - `ResolvePresets(excelHeaders, currentMap, fieldNames)` → 返回 (行表头, 行预设) 列表,重复字段只留第一个。
  - `AvailableFields(fieldNames, usedFields, currentValue)` → 可用字段列表。
- 现有 63 个测试保持通过。

## 兼容性

- 单字段正常映射行为不变。
- 旧的重复配置打开时自动清理(只留第一个),不破坏现有映射。
- 保存后生成的字段映射全局唯一,避免 INSERT 重复列问题。
