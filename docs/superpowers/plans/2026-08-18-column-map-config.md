# ESQLNew 列映射配置化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将写死在 ColumnMapper 的中文表头→英文字段映射抽到外部 JSON 配置,按目标表组织,支持多表通用导入。

**Architecture:** 新增 `Core/ColumnMapStore` 负责读写 `%APPDATA%\ESQLNew\column-maps.json`(键=表名,值=映射字典),内置 81 条维修单映射作为迁移源;`ColumnMapper.Map` 增加接受外部映射的重载,未命中退回名称匹配;`ImportEngine.Run` 按目标表加载映射传入比对。

**Tech Stack:** .NET Framework 4.6.2, Newtonsoft.Json 13.0.3, xUnit 2.4.2。

## Global Constraints

- 目标框架 net462,LangVersion 7.3,代码不写任何注释,UI 中文。
- 构建:`dotnet build ESQLNew.sln` 0 错误;测试:`dotnet test ESQLNew.sln` 全绿(当前 42 个)。
- 配置文件路径:`%APPDATA%\ESQLNew\column-maps.json`。
- 映射结构:顶层键=表名,值=`{ 中文表头: 英文字段 }` 字典。
- 比对优先级:该表映射 → 名称直接比对(大小写不敏感)→ 未匹配。
- 配置文件缺失时用内置 81 条维修单映射初始化并保存;损坏时回退内置映射。
- 分支 `feature/column-map-config`,不 push(除非用户要求)。
- 目标表由用户手动在数据库创建,程序不建表。

---

### Task 1: ColumnMapStore 配置读写 + 内置映射迁移

**Files:**
- Create: `src/ESQLNew/Core/ColumnMapStore.cs`
- Test: `tests/ESQLNew.Tests/ColumnMapStoreTests.cs`

**Interfaces:**
- Produces:
  - `static class ESQLNew.Core.ColumnMapStore`
    - `static string ConfigPath` — `Path.Combine(Environment.SpecialFolder.ApplicationData, "ESQLNew", "column-maps.json")`
    - `static Dictionary<string, Dictionary<string, string>> BuiltinMap` — 当前 ColumnMapper 的 81 条维修单映射(键 `sx_qms_repair_info_history`)
    - `static Dictionary<string, Dictionary<string, string>> Load(string path)` — 读取 JSON;文件不存在→用 BuiltinMap 初始化并 Save 后返回;解析失败→返回 BuiltinMap 包装
    - `static void Save(string path, Dictionary<string, Dictionary<string, string>> maps)` — 序列化写回,Formatting.Indented,自动建目录
    - `static Dictionary<string, string> GetMap(string path, string table)` — 返回指定表映射,无则空字典

- [ ] **Step 1: 写失败的测试**

`tests/ESQLNew.Tests/ColumnMapStoreTests.cs` 新建:

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class ColumnMapStoreTests
    {
        private static string TempPath()
        {
            return Path.Combine(Path.GetTempPath(), "colmap_" + Guid.NewGuid().ToString("N") + ".json");
        }

        [Fact]
        public void BuiltinMap_ContainsRepairTableWith81Entries()
        {
            var map = ColumnMapStore.BuiltinMap;
            Assert.True(map.ContainsKey("sx_qms_repair_info_history"));
            Assert.Equal(81, map["sx_qms_repair_info_history"].Count);
            Assert.Equal("repair_order_no", map["sx_qms_repair_info_history"]["维修单号"]);
        }

        [Fact]
        public void SaveThenLoad_RoundTrip()
        {
            var path = TempPath();
            try
            {
                var maps = new Dictionary<string, Dictionary<string, string>>
                {
                    { "my_table", new Dictionary<string, string> { { "订单号", "order_no" } } }
                };
                ColumnMapStore.Save(path, maps);
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("my_table"));
                Assert.Equal("order_no", loaded["my_table"]["订单号"]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_MissingFile_InitializesWithBuiltin()
        {
            var path = TempPath();
            try
            {
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
                Assert.True(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_CorruptFile_FallsBackToBuiltin()
        {
            var path = TempPath();
            try
            {
                File.WriteAllText(path, "{ not valid json");
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void GetMap_MissingTable_ReturnsEmpty()
        {
            var path = TempPath();
            try
            {
                ColumnMapStore.Save(path, new Dictionary<string, Dictionary<string, string>>());
                var map = ColumnMapStore.GetMap(path, "no_such_table");
                Assert.NotNull(map);
                Assert.Empty(map);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
```

- [ ] **Step 2: 跑测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ColumnMapStoreTests"`
Expected: FAIL(ColumnMapStore 未定义)。

- [ ] **Step 3: 实现 ColumnMapStore**

创建 `src/ESQLNew/Core/ColumnMapStore.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace ESQLNew.Core
{
    public static class ColumnMapStore
    {
        public static string ConfigPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ESQLNew",
                    "column-maps.json");
            }
        }

        public static Dictionary<string, Dictionary<string, string>> BuiltinMap
        {
            get
            {
                return new Dictionary<string, Dictionary<string, string>>
                {
                    { "sx_qms_repair_info_history", new Dictionary<string, string>
                        {
                            { "序号", "seq_no" },
                            { "担当", "person_in_charge" },
                            { "维修单号", "repair_order_no" },
                            { "次数", "order_times" },
                            { "重复单", "is_duplicate" },
                            { "地区", "region" },
                            { "省份", "province" },
                            { "城市", "city" },
                            { "大小型机", "machine_size" },
                            { "店铺编号", "shop_code" },
                            { "店铺名称", "shop_name" },
                            { "分类", "category" },
                            { "系列", "series" },
                            { "室内机型", "indoor_model" },
                            { "室外机型", "outdoor_model" },
                            { "机型", "model" },
                            { "机型年度", "model_year" },
                            { "机型年份", "year_label" },
                            { "室内机编号", "indoor_serial_no" },
                            { "室外机编号", "outdoor_serial_no" },
                            { "购买日期", "purchase_date" },
                            { "安装日期", "installation_date" },
                            { "故障日期", "failure_date" },
                            { "使用期限（天）", "service_life_days" },
                            { "年限", "age_limit" },
                            { "销售店", "sales_shop" },
                            { "安装店编码", "installation_shop_code" },
                            { "安装店", "installation_shop" },
                            { "受理时间", "acceptance_time" },
                            { "指定上门时间", "scheduled_visit_time" },
                            { "实际上门时间", "actual_visit_time" },
                            { "完成时间", "completion_time" },
                            { "首次维修日期", "first_repair_date" },
                            { "上次维修日期", "last_repair_date" },
                            { "维修次数", "repair_times" },
                            { "回访服务态度", "follow_up_attitude" },
                            { "回访满意度", "follow_up_satisfaction" },
                            { "其它意见", "other_opinion" },
                            { "担当审核", "person_in_charge_review" },
                            { "担当备注", "person_in_charge_remark" },
                            { "担当审核日期", "person_in_charge_review_date" },
                            { "回访", "follow_up" },
                            { "回访员", "follow_up_staff" },
                            { "回访日期", "follow_up_date" },
                            { "回访备注", "follow_up_remark" },
                            { "部长审核", "department_head_review" },
                            { "部长备注", "department_head_remark" },
                            { "部长审核日期", "department_head_review_date" },
                            { "单据状态", "document_status" },
                            { "交通距离", "travel_distance" },
                            { "上传月", "upload_month" },
                            { "上传时间", "upload_time" },
                            { "店铺备注", "shop_remark" },
                            { "故障现象", "failure_phenomenon" },
                            { "故障症状", "failure_symptom" },
                            { "故障原因", "failure_cause" },
                            { "处理方法", "handling_method" },
                            { "部件编码", "component_code" },
                            { "部件名称", "component_name" },
                            { "部件价格", "component_price" },
                            { "是否返厂机", "is_return_factory_machine" },
                            { "受理单号", "acceptance_no" },
                            { "是否换件", "is_replaced" },
                            { "转录保修单", "transfer_warranty" },
                            { "结算状态", "settlement_status" },
                            { "结算期数", "settlement_period" },
                            { "基本维修费", "basic_repair_fee" },
                            { "交通费", "transportation_fee" },
                            { "速度费", "speed_fee" },
                            { "信息费", "information_fee" },
                            { "上门费", "visit_fee" },
                            { "完成费", "completion_fee" },
                            { "扣罚款", "penalty_amount" },
                            { "扣罚款原因", "penalty_reason" },
                            { "调整费用", "adjustment_fee" },
                            { "调整备注", "adjustment_remark" },
                            { "结算金额", "settlement_amount" },
                            { "合计", "total_amount" },
                            { "项目", "item" },
                            { "室内外区分", "indoor_outdoor_type" },
                            { "期限", "warranty_period" }
                        } }
                };
            }
        }

        public static Dictionary<string, Dictionary<string, string>> Load(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    var init = BuiltinMap;
                    Save(path, init);
                    return init;
                }
                var text = File.ReadAllText(path);
                var parsed = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(text);
                if (parsed != null) return parsed;
                return BuiltinMap;
            }
            catch
            {
                return BuiltinMap;
            }
        }

        public static void Save(string path, Dictionary<string, Dictionary<string, string>> maps)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(path, JsonConvert.SerializeObject(maps, Formatting.Indented));
        }

        public static Dictionary<string, string> GetMap(string path, string table)
        {
            var maps = Load(path);
            Dictionary<string, string> result;
            if (maps.TryGetValue(table, out result))
                return result;
            return new Dictionary<string, string>();
        }
    }
}
```

- [ ] **Step 4: 跑测试确认通过**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ColumnMapStoreTests"`
Expected: 5 个测试全过。

- [ ] **Step 5: 提交**

```bash
git add src/ESQLNew/Core/ColumnMapStore.cs tests/ESQLNew.Tests/ColumnMapStoreTests.cs
git commit -m "feat: add column map store with builtin repair-table mapping"
```

---

### Task 2: ColumnMapper 支持外部映射

**Files:**
- Modify: `src/ESQLNew/Import/ColumnMapper.cs`(删除 FieldMap 字段,新增 Map 3 参重载)
- Test: `tests/ESQLNew.Tests/ColumnMapperTests.cs`

**Interfaces:**
- Consumes: `ColumnMapStore`(Task 1)。
- Produces:
  - `Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)` — 保留,内部调用 3 参版传 `ColumnMapStore.BuiltinMap["sx_qms_repair_info_history"]`
  - `Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns, IDictionary<string,string> fieldMap)` — 优先 fieldMap,未命中退回名称比对
- 注意:现有 ColumnMapperTests 的 `Map_MatchesByName_CaseInsensitive` 期望表头"担当"匹配字段"担当"——3 参版在 fieldMap 有"担当"→"person_in_charge"时,会优先映射到 person_in_charge,但该测试的字段列表只有"担当"没有 person_in_charge,所以映射查找失败后应退回名称比对命中"担当"。此逻辑须保持,测试通过即验证。

- [ ] **Step 1: 写失败的测试**

`tests/ESQLNew.Tests/ColumnMapperTests.cs` 追加:

```csharp
        [Fact]
        public void Map_ExternalMapTakesPriority()
        {
            var headers = new List<string> { "维修单号" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var fieldMap = new Dictionary<string, string> { { "维修单号", "repair_order_no" } };
            var map = ColumnMapper.Map(headers, cols, fieldMap);
            Assert.True(map[0].Matched);
            Assert.Equal("repair_order_no", map[0].TableField);
        }

        [Fact]
        public void Map_ExternalMapMiss_FallsBackToName()
        {
            var headers = new List<string> { "自定义列" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "自定义列", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var fieldMap = new Dictionary<string, string> { { "维修单号", "repair_order_no" } };
            var map = ColumnMapper.Map(headers, cols, fieldMap);
            Assert.True(map[0].Matched);
            Assert.Equal("自定义列", map[0].TableField);
        }
```

- [ ] **Step 2: 跑测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ColumnMapperTests"`
Expected: 2 个新测试失败(3 参 Map 不存在)。

- [ ] **Step 3: 实现**

修改 `src/ESQLNew/Import/ColumnMapper.cs`:
1. 删除 `FieldMap` 字段(第 9-92 行,含 `private static readonly Dictionary<string, string> FieldMap = ...` 整块)。
2. 原 `Map(2参)` 方法体改为:

```csharp
        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)
        {
            Dictionary<string, string> builtin;
            ColumnMapStore.BuiltinMap.TryGetValue("sx_qms_repair_info_history", out builtin);
            return Map(excelHeaders, tableColumns, builtin ?? new Dictionary<string, string>());
        }

        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns,
            IDictionary<string, string> fieldMap)
        {
            var result = new List<ColumnMapping>(excelHeaders.Count);
            foreach (var h in excelHeaders)
            {
                ColumnInfo matched = null;
                string mappedField = null;
                string key = h.Trim();
                if (fieldMap != null && fieldMap.TryGetValue(key, out mappedField))
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), mappedField, StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                if (matched == null)
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), h.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                result.Add(new ColumnMapping
                {
                    ExcelColumn = h,
                    TableField = matched != null ? matched.Name : null,
                    Matched = matched != null,
                    TableColumn = matched
                });
            }
            return result;
        }
```

需在文件顶部加 `using ESQLNew.Core;`。

- [ ] **Step 4: 跑测试确认通过**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`
Expected: 全部通过(42 + 2 = 44)。

- [ ] **Step 5: 提交**

```bash
git add src/ESQLNew/Import/ColumnMapper.cs tests/ESQLNew.Tests/ColumnMapperTests.cs
git commit -m "feat: support external column map in ColumnMapper"
```

---

### Task 3: ImportEngine 与 UI 接入外部映射

**Files:**
- Modify: `src/ESQLNew/Import/ImportEngine.cs`
- Modify: `src/ESQLNew/MainForm.Tabs.cs`
- Test: `tests/ESQLNew.Tests/ImportEngineTests.cs`

**Interfaces:**
- Consumes: `ColumnMapStore.Load/GetMap`(Task 1)、`ColumnMapper.Map 3参`(Task 2)。
- Produces: 无新公开 API;行为变化——导入/预览按目标表加载外部映射。

- [ ] **Step 1: 写失败的测试**

`tests/ESQLNew.Tests/ImportEngineTests.cs` 追加:

```csharp
        [Fact]
        public void Run_UsesTableSpecificExternalMap()
        {
            var path = Path.Combine(Path.GetTempPath(), "impmap_" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                ColumnMapStore.Save(path, new Dictionary<string, Dictionary<string, string>>
                {
                    { "custom_table", new Dictionary<string, string> { { "自定义表头", "custom_field" } } }
                });
                var map = ColumnMapStore.GetMap(path, "custom_table");
                Assert.Equal("custom_field", map["自定义表头"]);
            }
            finally
            {
                File.Delete(path);
            }
        }
```

需要在文件顶部加 `using System.IO;` 与 `using ESQLNew.Core;`。此测试验证 GetMap 取表专属映射——为 ImportEngine.Run 使用外部映射提供依据。

- [ ] **Step 2: 跑测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~ImportEngineTests"`
Expected: 新增测试失败(ColumnMapStore 不可见于该命名空间或未引用——实际因 using 缺失编译失败)。

- [ ] **Step 3: 实现 ImportEngine 接入**

修改 `src/ESQLNew/Import/ImportEngine.cs`:
1. 文件顶部加 `using ESQLNew.Core;`
2. `Run` 方法中,`var mappings = ColumnMapper.Map(headers, cols);` 改为:

```csharp
            var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
            var mappings = ColumnMapper.Map(headers, cols, fieldMap);
```

- [ ] **Step 4: 实现 UI 接入**

修改 `src/ESQLNew/MainForm.Tabs.cs`:
1. `Preview_Click` 中 `var mappings = ColumnMapper.Map(headers, cols);` 改为:

```csharp
                var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
                var mappings = ColumnMapper.Map(headers, cols, fieldMap);
```

- [ ] **Step 5: 跑测试确认通过**

Run: `dotnet build ESQLNew.sln`(0 错误)然后 `dotnet test ESQLNew.sln`
Expected: 全部通过(45)。

- [ ] **Step 6: 提交**

```bash
git add src/ESQLNew/Import/ImportEngine.cs src/ESQLNew/MainForm.Tabs.cs tests/ESQLNew.Tests/ImportEngineTests.cs
git commit -m "feat: load per-table column map from config for import and preview"
```

---

### Task 4: 收尾验证

**Files:**
- 无代码改动;验证与文档。

- [ ] **Step 1: 全量构建与测试**

Run: `dotnet build ESQLNew.sln` 与 `dotnet test ESQLNew.sln`
Expected: 0 错误,45 测试全绿。

- [ ] **Step 2: 验证配置文件生成**

Run: `Test-Path "$env:APPDATA\ESQLNew\column-maps.json"`
说明:首次运行程序或测试后生成,含 `sx_qms_repair_info_history` 的 81 条。若无,手动运行 `dotnet run --project src/ESQLNew` 一次再查。

- [ ] **Step 3: 代码审查**

按 superpowers:requesting-code-review 派发本分支全量 review(范围 = 本功能 3 个实现提交 + 文档提交的 diff)。

- [ ] **Step 4: 提交审查修复**

若审查发现问题,修复并提交 `fix: address column map config review findings`。

---

## Self-Review

**Spec coverage:**
- 配置路径与结构:Task 1 ✓
- BuiltinMap 81 条迁移:Task 1 ✓
- Map 3 参重载 + 退回名称比对:Task 2 ✓
- ImportEngine 按表加载:Task 3 ✓
- UI 预览接入:Task 3 ✓
- 错误处理(缺失/损坏回退):Task 1 Load ✓
- 测试:Task 1-3 ✓
- 验收 1(现有行为不变):2 参 Map 内部走 BuiltinMap 维修单表映射 ✓
- 验收 2(新表配置即用):GetMap 按表取映射 ✓
- 验收 3(缺失/损坏不崩溃):Load 回退 ✓

**Placeholder scan:** 无 TBD/TODO;所有步骤含完整代码。81 条映射在 Task 1 完整列出,无省略。

**Type consistency:** `ColumnMapStore.Load/Save/GetMap/ConfigPath/BuiltinMap` 在 Task 1 定义,Task 2/3 使用一致;`ColumnMapper.Map 3参` 签名在 Task 2 定义,Task 3 使用一致;`Dictionary<string, Dictionary<string, string>>` 类型贯通全计划。