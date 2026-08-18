# ESQLNew 界面编辑列映射设计

日期:2026-08-18
状态:待审阅

## 背景

列映射配置化后,新表导入需在 `column-maps.json` 手动添加映射,无界面支持。用户在导入时遇到"表不在配置中→全部未匹配",希望能在界面直接添加/编辑映射。

## 需求(经逐项澄清确认)

1. 在导入 Tab"字段映射"预览处提供"编辑映射"按钮,弹出对话框逐列编辑。
2. 弹窗每行 = Excel 表头 + 英文字段下拉;英文字段候选从当前目标表真实字段(GetTableColumns)拉取,未连接则无候选。
3. 未匹配的行默认"不映射"(下拉置空),用户自行选择英文字段。
4. 保存只针对当前目标表:新表新增一段映射,已有表覆盖该表映射,不影响其它表。
5. 保存后配置立即生效,可重新预览/导入。

## 方案:弹窗逐列编辑

### 交互流程

```
导入 Tab:选择文件 + 目标表 → 匹配预览 → 点「编辑映射」
  │
  ▼ 弹出「编辑列映射」对话框
  │  · 标题显示目标表名
  │  · 逐行:Excel 表头(只读) | 英文字段下拉(可空)
  │  · 英文字段候选 = GetTableColumns(当前连接, 目标表) 的字段名
  │  · 初始值 = 当前配置中该表的映射(已匹配的列预选其字段)
  │  · 未匹配列默认下拉为空
  │
  ▼ 用户编辑 → 点「保存」
  │  · 只取下拉非空的那些行,生成 表头→字段 映射
  │  · ColumnMapStore.Save 写回配置文件(新增/覆盖该表)
  │  · 关闭弹窗
  │
  ▼ 重新执行匹配预览(用新映射)→ 应显示匹配成功
```

### 组件改动

**新增 `src/ESQLNew/MainForm.MappingDialog.cs`(partial 类的一部分):**
- `class ColumnMapDialog : Form` — 弹窗
  - 构造参数:`string table`, `IList<string> excelHeaders`, `IList<ColumnInfo> tableColumns`, `Dictionary<string,string> currentMap`
  - 界面:DataGridView(列:Excel表头只读 / 英文字段下拉组合框列)
  - `Dictionary<string,string> Result` — 保存后的映射(表头→字段)
  - 「保存」按钮:收集下拉非空行生成 Result,DialogResult.OK
  - 「取消」按钮:DialogResult.Cancel

**修改 `src/ESQLNew/MainForm.Tabs.cs`:**
- 导入 Tab 加「编辑映射」按钮(放在匹配预览按钮旁或字段映射页签内)
- 点击处理:
  1. 校验已选文件与目标表(同预览)
  2. 读 Excel 表头、GetTableColumns(当前连接, 目标表)
  3. 读当前配置 `ColumnMapStore.GetMap(ConfigPath, table)`
  4. 打开弹窗;保存后提示成功,并自动重跑匹配预览显示新结果

### 关键细节

- 英文字段下拉列:DataGridViewComboBoxColumn,Items = tableColumns 字段名,可清空(空 = 不映射)。用 DataGridView 的编辑控件支持用户点下拉选择;允许空值。
- 配置写入前需确认表存在映射条目;无则新增。不影响其它表(Load→改当前表→Save 整体写回)。
- 保存后无需重启,后续预览/导入直接用新映射。
- 若未连接(GetTableColumns 抛异常),弹窗仍可打开但英文字段下拉为空,提示"无法读取目标表字段,请先测试连接"。

### 测试

- 弹窗为 UI,不做单元测试;核心逻辑(从弹窗结果构造映射字典、保存)可抽取为 `ColumnMapStore` 已有方法或辅助方法测试。
- 新增测试:验证"下拉非空行→映射字典"的收集逻辑(若有可测辅助方法)、`ColumnMapStore` 已有 Save/Load 覆盖。
- 既有 51 测试保持全绿。

## 验收标准

1. 导入 Tab 出现「编辑映射」按钮。
2. 点击弹出对话框,逐行显示 Excel 表头 + 英文字段下拉,候选来自目标表字段。
3. 未匹配列默认下拉为空;已匹配列预选当前配置字段。
4. 保存后配置文件新增/更新当前表映射,其它表不受影响。
5. 保存后重跑预览,未匹配列按新映射匹配成功。
6. 构建 0 错误,测试全绿。