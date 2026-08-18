# ESQLNew 列映射配置化设计

日期:2026-08-18
状态:待审阅

## 背景

当前 `ColumnMapper.cs` 内置 81 条"中文表头 → 英文字段"映射,硬编码、只适用维修单表。换新业务(新 Excel 结构 + 新目标表)需改代码重新编译,无法通用。

## 需求

1. 将列映射从代码抽取到外部配置文件,按目标表组织,多表互不干扰。
2. 换新业务时:数据库建好表 + 配置加一段映射 → 直接导入,不改代码不重编译。
3. 现有 81 条映射首次运行时自动迁移进配置,现有功能不回归。
4. 目标表由用户手动在数据库创建(不自动建表)。

## 方案:映射配置化

### 配置文件

路径:`%APPDATA%\ESQLNew\column-maps.json`

结构:

```json
{
  "sx_qms_repair_info_history": {
    "维修单号": "repair_order_no",
    "担当": "person_in_charge",
    "结算金额": "settlement_amount"
  }
}
```

键 = 目标表名,值 = 该表的中文表头 → 英文字段映射字典。

### 组件改动

**新增 `src/ESQLNew/Core/ColumnMapStore.cs`:**
- `static string ConfigPath` — 指向 `%APPDATA%\ESQLNew\column-maps.json`
- `static Dictionary<string, Dictionary<string, string>> Load(string path)` — 读取 JSON;文件缺失/损坏时返回空字典并尝试用内置默认映射初始化
- `static void Save(string path, Dictionary<...>)` — 写回 JSON(格式化缩进)
- `static Dictionary<string, string> GetMap(string path, string table)` — 返回指定表的映射,无则空字典
- `static Dictionary<string, string> BuiltinMap` — 现有 81 条内置映射(迁移源)

**修改 `src/ESQLNew/Import/ColumnMapper.cs`:**
- `Map` 增加重载 `Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns, IDictionary<string,string> fieldMap)` — 优先用外部映射,未命中退回名称匹配
- 原 `Map(2参)` 保持,内部调用 3 参版(传 `ColumnMapStore.BuiltinMap`)以兼容既有测试
- `BuiltinMap` 的 81 条常量从 ColumnMapper 迁到 ColumnMapStore(单一数据源)

**修改 `src/ESQLNew/Import/ImportEngine.cs`:**
- `Run` 中调用 `ColumnMapStore.Load` 读取配置,`GetMap(table)` 取该表映射,传入 `ColumnMapper.Map`

**修改 `src/ESQLNew/MainForm.Tabs.cs`:**
- 预览/导入前读取配置中的表映射并用于比对
- (可选,本版本不做)界面编辑映射按钮 — 留待后续

### 数据流

```
导入 Tab 选目标表
  → ColumnMapStore.Load → GetMap(table) → 该表中文→英文映射
  → ColumnMapper.Map(headers, cols, fieldMap)
      → ① 查该表映射  ② 未命中→名称比对  ③ 都没有→未匹配
  → 预览 / 导入
```

### 迁移策略

首次运行 `Load` 时:若配置文件不存在,用 `BuiltinMap` 初始化配置并保存(含维修单表 81 条)。之后以配置文件为准;内置映射仅作为初始化源和没有配置文件时的兜底。

### 错误处理

- 配置文件损坏(JSON 解析失败):回退内置映射,提示"列映射配置读取失败,已使用内置映射"。
- 表在配置中无映射:空字典,走名称比对兜底(可能 0 匹配,预览可见)。

### 测试

- `ColumnMapStore` 单元测试:Load/Save/GetMap 往返、文件缺失返回空、损坏 JSON 回退、BuiltinMap 含 81 条维修单映射。
- `ColumnMapper` 新增测试:外部映射优先、外部映射未命中退回名称匹配。
- 既有 42 测试保持全绿。

## 验收标准

1. 现有维修单表导入行为不变(81 条迁移进配置后仍 81/81 匹配)。
2. 配置文件中新增一个表的映射,可导入该表(手动建表后)。
3. 配置缺失/损坏时不崩溃,回退内置映射。
4. 构建 0 错误,全部测试通过。
- 注意:内置映射仅对表名 `sx_qms_repair_info_history` 生效;同结构不同表名需在配置文件为其添加映射,否则按名称比对可能 0 匹配。