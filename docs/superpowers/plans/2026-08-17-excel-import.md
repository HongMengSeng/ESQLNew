# Excel 批量导入 MySQL 工具 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建 .NET Framework 4.6.2 WinForms 桌面工具,支持动态 MySQL 连接配置(表单 + 完整连接串 + 测试连接)、Excel(.xlsx/.xls)流式读取、列名匹配预览、多值 INSERT 批量导入(1万~100万行)、可视化进度 + 汇总报告,以及可选 SQLite 日志。

**Architecture:** WinForms 单窗口三 Tab(连接/导入/日志),核心逻辑抽成可单测的服务类,后台线程导入经 `BeginInvoke` 刷新 UI。SDK 风格 csproj,`TargetFramework=net462`,`UseWindowsForms=true`,本地缺失的 4.6.2 目标程序集由 NuGet `Microsoft.NETFramework.ReferenceAssemblies.net462` 补齐。

**Tech Stack:** C# / .NET Framework 4.6.2 / WinForms;MySql.Data 8.0.33;ExcelDataReader 3.6.0 + ExcelDataReader.DataSet;System.Data.SQLite 1.0.118;Newtonsoft.Json 13.0.3;xUnit + Microsoft.NET.Test.Sdk

## Global Constraints

- 目标框架:`net462`(禁止改低);仅支持 Windows
- 数据库:远端 MySQL/MariaDB `192.168.201.112:3306`(连接配置界面可改)
- Excel 格式:仅 `.xlsx` / `.xls`;列名匹配表字段,逐行容错继续
- 包版本固定:MySql.Data=8.0.33, ExcelDataReader=3.6.0, ExcelDataReader.DataSet=3.6.0, System.Data.SQLite=1.0.118, Newtonsoft.Json=13.0.3
- 所有代码不写注释
- 构建:`dotnet build ESQLNew.sln`;测试:`dotnet test ESQLNew.sln`
- git 仓库 `origin` 已指向 GitHub HongMengSeng/ESQLNew

## 模板数据结构(参照,用于 ConvertValue 与性能实测)

- 模板:`历史维修安装数据.xlsx`,sheet 名 `SxQmsRepairInfoHistory`,80 列(A~CB)
- 类型分布:文本/编码类约 35 列(担当/维修单号/机型/店铺名称等)、数值类约 18 列(次数/使用期限(天)/基本维修费/结算金额等)、日期时间类约 14 列(购买日期/安装日期/受理时间/完成时间等)
- 目标表已存在,字段与 80 列表头同名;表名在界面手动输入,选文件后自动以第一个 sheet 名预填

---

### Task 1: 工程骨架 + NuGet 还原 + 空壳可构建

**Files:**
- Create: `ESQLNew.sln`
- Create: `src/ESQLNew/ESQLNew.csproj`
- Create: `src/ESQLNew/Program.cs`
- Create: `src/ESQLNew/MainForm.cs`
- Create: `tests/ESQLNew.Tests/ESQLNew.Tests.csproj`
- Create: `tests/ESQLNew.Tests/PlaceholderTests.cs`

**Interfaces:**
- Consumes: 无
- Produces: 可构建的解决方案、可运行的 WinForms 空窗口、可运行的 xUnit 测试项目

- [ ] **Step 1**: 创建 SDK 风格 csproj

`src/ESQLNew/ESQLNew.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net462</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <LangVersion>7.3</LangVersion>
    <AssemblyName>ESQLNew</AssemblyName>
    <RootNamespace>ESQLNew</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies.net462" Version="1.0.3" PrivateAssets="All" />
    <PackageReference Include="MySql.Data" Version="8.0.33" />
    <PackageReference Include="ExcelDataReader" Version="3.6.0" />
    <PackageReference Include="ExcelDataReader.DataSet" Version="3.6.0" />
    <PackageReference Include="System.Data.SQLite" Version="1.0.118" />
    <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
  </ItemGroup>
</Project>
```

`src/ESQLNew/Program.cs`:
```csharp
using System;
using System.Windows.Forms;

namespace ESQLNew
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
```

`src/ESQLNew/MainForm.cs`(暂空,后续任务填充):
```csharp
using System.Windows.Forms;

namespace ESQLNew
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            Text = "ESQLNew - Excel 批量导入";
        }
    }
}
```

- [ ] **Step 2**: 创建测试项目

`tests/ESQLNew.Tests/ESQLNew.Tests.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net462</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies.net462" Version="1.0.3" PrivateAssets="All" />
    <PackageReference Include="xunit" Version="2.4.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.4.5" />
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.5.0" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\..\src\ESQLNew\ESQLNew.csproj" />
  </ItemGroup>
</Project>
```

`tests/ESQLNew.Tests/PlaceholderTests.cs`:
```csharp
using Xunit;

namespace ESQLNew.Tests
{
    public class PlaceholderTests
    {
        [Fact]
        public void Sanity_True()
        {
            Assert.True(true);
        }
    }
}
```

- [ ] **Step 3**: 创建解决方案并还原构建

```bash
dotnet new sln -n ESQLNew
dotnet sln add src/ESQLNew/ESQLNew.csproj tests/ESQLNew.Tests/ESQLNew.Tests.csproj
dotnet restore
dotnet build ESQLNew.sln
```
Expected: `Build succeeded`;确认从 nuget.org 还原成功(ReferenceAssemblies、MySql.Data 等)。

- [ ] **Step 4**: 运行测试

```bash
dotnet test ESQLNew.sln
```
Expected: 1 个测试通过。

- [ ] **Step 5**: 提交

```bash
git add -A
git commit -m "feat: scaffold net462 winforms solution"
```

---

### Task 2: AppConfig + MySqlConnectionBuilder

**Files:**
- Create: `src/ESQLNew/Core/MySqlConnectionBuilder.cs`
- Create: `src/ESQLNew/Core/AppConfig.cs`
- Create: `tests/ESQLNew.Tests/MySqlConnectionBuilderTests.cs`
- Create: `tests/ESQLNew.Tests/AppConfigTests.cs`

**Interfaces:**
- Consumes: Newtonsoft.Json
- Produces:
  - `static string MySqlConnectionBuilder.Build(string server, int port, string user, string password, string database)`
  - `AppConfig { string Server; int Port; string User; string Password; string Database; string RawConnectionString; bool UseRaw; int BatchSize; int CommitEvery; bool KeepLogs; int LogRetentionDays; }` 及 `static AppConfig Load(string path)`, `void Save(string path)`, 默认值 `Port=3306, BatchSize=2000, CommitEvery=5000, KeepLogs=true, LogRetentionDays=30`

- [ ] **Step 1**: 写失败测试 `MySqlConnectionBuilderTests`

```csharp
using Xunit;

namespace ESQLNew.Tests
{
    public class MySqlConnectionBuilderTests
    {
        [Fact]
        public void Build_IncludesAllParts()
        {
            var cs = ESQLNew.Core.MySqlConnectionBuilder.Build(
                "192.168.201.112", 3306, "root", "pwd", "testdb");
            Assert.Contains("Server=192.168.201.112", cs);
            Assert.Contains("Port=3306", cs);
            Assert.Contains("Uid=root", cs);
            Assert.Contains("Pwd=pwd", cs);
            Assert.Contains("Database=testdb", cs);
            Assert.Contains("SslMode=None", cs);
            Assert.Contains("AllowLoadLocalInfile=false", cs);
            Assert.Contains("CharSet=utf8mb4", cs);
        }
    }
}
```

- [ ] **Step 2**: 运行确认失败(`dotnet test`,缺 MySqlConnectionBuilder → 编译错误即视为失败)
- [ ] **Step 3**: 实现 `MySqlConnectionBuilder.cs`

```csharp
namespace ESQLNew.Core
{
    public static class MySqlConnectionBuilder
    {
        public static string Build(string server, int port, string user, string password, string database)
        {
            return string.Format(
                "Server={0};Port={1};Uid={2};Pwd={3};Database={4};SslMode=None;AllowLoadLocalInfile=false;CharSet=utf8mb4;",
                server, port, user, password, database);
        }
    }
}
```

- [ ] **Step 4**: 写 `AppConfigTests`(临时目录 JSON 往返)

```csharp
using System;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class AppConfigTests
    {
        [Fact]
        public void SaveLoad_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = new AppConfig
            {
                Server = "192.168.201.112",
                Port = 3306,
                User = "root",
                Password = "pwd",
                Database = "testdb",
                RawConnectionString = "",
                UseRaw = false,
                BatchSize = 2000,
                CommitEvery = 5000,
                KeepLogs = true,
                LogRetentionDays = 30
            };
            cfg.Save(path);
            var loaded = AppConfig.Load(path);
            Assert.Equal("192.168.201.112", loaded.Server);
            Assert.Equal(3306, loaded.Port);
            Assert.Equal("root", loaded.User);
            Assert.Equal("pwd", loaded.Password);
            Assert.Equal("testdb", loaded.Database);
            Assert.Equal(2000, loaded.BatchSize);
            Assert.Equal(true, loaded.KeepLogs);
            File.Delete(path);
        }

        [Fact]
        public void Load_Missing_ReturnsDefaults()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_missing_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = AppConfig.Load(path);
            Assert.Equal(3306, cfg.Port);
            Assert.Equal(2000, cfg.BatchSize);
            Assert.Equal(5000, cfg.CommitEvery);
            Assert.True(cfg.KeepLogs);
        }
    }
}
```

- [ ] **Step 5**: 实现 `AppConfig.cs`

```csharp
using System;
using System.IO;
using Newtonsoft.Json;

namespace ESQLNew.Core
{
    public class AppConfig
    {
        public string Server { get; set; }
        public int Port { get; set; } = 3306;
        public string User { get; set; }
        public string Password { get; set; }
        public string Database { get; set; }
        public string RawConnectionString { get; set; }
        public bool UseRaw { get; set; }
        public int BatchSize { get; set; } = 2000;
        public int CommitEvery { get; set; } = 5000;
        public bool KeepLogs { get; set; } = true;
        public int LogRetentionDays { get; set; } = 30;

        public static AppConfig Load(string path)
        {
            try
            {
                if (File.Exists(path))
                    return JsonConvert.DeserializeObject<AppConfig>(File.ReadAllText(path));
            }
            catch
            {
            }
            return new AppConfig();
        }

        public void Save(string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
```

- [ ] **Step 6**: `dotnet test ESQLNew.sln` 全绿
- [ ] **Step 7**: 提交 `feat: connection builder + app config`

---

### Task 3: Excel 流式读取(ExcelStreamReader)

**Files:**
- Create: `src/ESQLNew/Excel/ExcelStreamReader.cs`
- Create: `tests/ESQLNew.Tests/ExcelStreamReaderTests.cs`

**Interfaces:**
- Produces:
  - `static IList<string> SheetNames(string path)`
  - `static IList<string> ReadHeaders(string path, string sheetName = null)`(取首个工作表)
  - `static IEnumerable<object[]> ReadRows(string path, string sheetName = null)`(流式,跳过表头,逐行产出)

- [ ] **Step 1**: 写失败测试(用模板真实文件做集成测试)

`tests/ESQLNew.Tests/ExcelStreamReaderTests.cs`:
```csharp
using System;
using System.IO;
using System.Linq;
using ESQLNew.Excel;
using Xunit;

namespace ESQLNew.Tests
{
    public class ExcelStreamReaderTests
    {
        private static string Template => @"C:\Users\Kuade\Downloads\历史维修安装数据.xlsx";

        [Fact]
        public void SheetNames_ReturnsTemplateSheet()
        {
            var names = ExcelStreamReader.SheetNames(Template);
            Assert.Contains("SxQmsRepairInfoHistory", names);
        }

        [Fact]
        public void ReadHeaders_Returns80Columns()
        {
            var headers = ExcelStreamReader.ReadHeaders(Template);
            Assert.Equal(80, headers.Count);
            Assert.Equal("担当", headers[0]);
            Assert.Equal("期限", headers[79]);
        }

        [Fact]
        public void ReadRows_StreamsRows()
        {
            var rows = ExcelStreamReader.ReadRows(Template).ToList();
            Assert.True(rows.Count >= 300, "模板应包含 347 行数据");
        }

        [Fact]
        public void MissingFile_Throws()
        {
            Assert.Throws<FileNotFoundException>(() => ExcelStreamReader.ReadHeaders(@"C:\does_not_exist.xlsx"));
        }
    }
}
```

注意:模板文件 deflate 流被截断(前 491520 字节可读),ReadRows 断言放宽为 >=300;若 ExcelDataReader 也抛异常,将断言改为捕获异常并记录,标记该文件损坏,由用户重新导出全量文件。

- [ ] **Step 2**: 运行确认失败(编译错误即失败)
- [ ] **Step 3**: 实现

`src/ESQLNew/Excel/ExcelStreamReader.cs`:
```csharp
using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ESQLNew.Excel
{
    public static class ExcelStreamReader
    {
        public static IList<string> SheetNames(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet();
                var names = new List<string>();
                foreach (System.Data.DataTable t in ds.Tables)
                    names.Add(t.TableName);
                return names;
            }
        }

        public static IList<string> ReadHeaders(string path, string sheetName = null)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    var name = reader.Name;
                    var header = new List<string>();
                    if (sheetName == null || string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                                header.Add(reader.GetValue(i) == null ? "" : reader.GetValue(i).ToString());
                        }
                        return header;
                    }
                } while (reader.NextResult());
            }
            return new List<string>();
        }

        public static IEnumerable<object[]> ReadRows(string path, string sheetName = null)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    var name = reader.Name;
                    if (sheetName == null || string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        bool first = true;
                        while (reader.Read())
                        {
                            if (first) { first = false; continue; }
                            var values = new object[reader.FieldCount];
                            for (int i = 0; i < reader.FieldCount; i++)
                                values[i] = reader.GetValue(i);
                            yield return values;
                        }
                        yield break;
                    }
                } while (reader.NextResult());
            }
        }
    }
}
```

- [ ] **Step 4**: `dotnet test ESQLNew.sln`;若模板文件损坏导致失败,记录到测试注释并确认异常内容
- [ ] **Step 5**: 提交 `feat: streaming excel reader`

---

### Task 4: 列匹配 + 值转换(ColumnMapper)

**Files:**
- Create: `src/ESQLNew/Import/ImportModels.cs`
- Create: `src/ESQLNew/Import/ColumnMapper.cs`
- Create: `tests/ESQLNew.Tests/ColumnMapperTests.cs`

**Interfaces:**
- Produces:
  - `ColumnInfo { string Name; string DataType; bool IsNullable; int MaxLength; }`
  - `ColumnMapping { string ExcelColumn; string TableField; bool Matched; }`
  - `static IList<ColumnMapping> ColumnMapper.Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)`(大小写不敏感)
  - `static object ColumnMapper.ConvertValue(object raw, ColumnInfo col)`(按 DATA_TYPE:char/varchar/text→string、int→long、decimal/double→decimal、datetime/timestamp→DateTime、空→null)

- [ ] **Step 1**: 写失败测试

`tests/ESQLNew.Tests/ColumnMapperTests.cs`:
```csharp
using System.Collections.Generic;
using ESQLNew.Import;
using Xunit;

namespace ESQLNew.Tests
{
    public class ColumnMapperTests
    {
        [Fact]
        public void Map_MatchesByName_CaseInsensitive()
        {
            var headers = new List<string> { "担当", "次数", "不存在的列" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "担当", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "次数", DataType = "int", IsNullable = false, MaxLength = 0 }
            };
            var map = ColumnMapper.Map(headers, cols);
            Assert.Equal(3, map.Count);
            Assert.True(map[0].Matched);
            Assert.Equal("担当", map[0].TableField);
            Assert.False(map[2].Matched);
        }

        [Fact]
        public void ConvertValue_EmptyToNull()
        {
            var col = new ColumnInfo { Name = "a", DataType = "varchar", IsNullable = true, MaxLength = 50 };
            Assert.Null(ColumnMapper.ConvertValue("", col));
        }

        [Fact]
        public void ConvertValue_DateOADateToDateTime()
        {
            var col = new ColumnInfo { Name = "d", DataType = "datetime", IsNullable = true, MaxLength = 0 };
            var dt = ColumnMapper.ConvertValue(new System.DateTime(2020, 4, 17), col);
            Assert.Equal(new System.DateTime(2020, 4, 17), dt);
        }

        [Fact]
        public void ConvertValue_NumericStringToLong()
        {
            var col = new ColumnInfo { Name = "n", DataType = "int", IsNullable = true, MaxLength = 0 };
            Assert.Equal(123L, ColumnMapper.ConvertValue("123", col));
        }
    }
}
```

- [ ] **Step 2**: 运行确认失败
- [ ] **Step 3**: 实现

`src/ESQLNew/Import/ImportModels.cs`:
```csharp
using System;

namespace ESQLNew.Import
{
    public class ColumnInfo
    {
        public string Name { get; set; }
        public string DataType { get; set; }
        public bool IsNullable { get; set; }
        public int MaxLength { get; set; }
    }

    public class ColumnMapping
    {
        public string ExcelColumn { get; set; }
        public string TableField { get; set; }
        public bool Matched { get; set; }
    }

    public class RowFailure
    {
        public long RowNumber { get; set; }
        public string Message { get; set; }
    }

    public class ImportResult
    {
        public long Total { get; set; }
        public long Succeeded { get; set; }
        public long Failed { get; set; }
        public System.Collections.Generic.List<RowFailure> Failures { get; set; } =
            new System.Collections.Generic.List<RowFailure>();
        public TimeSpan Elapsed { get; set; }
        public double RowsPerSecond { get; set; }
        public int BatchCount { get; set; }
    }

    public class ImportProgress
    {
        public long Processed { get; set; }
        public long Total { get; set; }
        public long Succeeded { get; set; }
        public long Failed { get; set; }
    }
}
```

`src/ESQLNew/Import/ColumnMapper.cs`:
```csharp
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ESQLNew.Import
{
    public static class ColumnMapper
    {
        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)
        {
            var result = new List<ColumnMapping>(excelHeaders.Count);
            foreach (var h in excelHeaders)
            {
                ColumnInfo matched = null;
                foreach (var c in tableColumns)
                {
                    if (string.Equals(c.Name.Trim('`'), h.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        matched = c;
                        break;
                    }
                }
                result.Add(new ColumnMapping
                {
                    ExcelColumn = h,
                    TableField = matched != null ? matched.Name : null,
                    Matched = matched != null
                });
            }
            return result;
        }

        public static object ConvertValue(object raw, ColumnInfo col)
        {
            if (raw == null) return null;
            string s = raw as string;
            if (s != null && s.Trim().Length == 0) return null;

            string type = (col.DataType ?? "").ToLowerInvariant();

            if (type == "datetime" || type == "timestamp" || type == "date" || type == "time")
            {
                var dt = raw as DateTime?;
                if (dt.HasValue) return dt.Value;
                if (s != null)
                {
                    DateTime parsed;
                    if (DateTime.TryParse(s.Trim(), out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析日期: " + s);
            }

            if (type == "int" || type == "bigint" || type == "smallint" || type == "tinyint" || type == "mediumint")
            {
                var n = raw as int?;
                if (n.HasValue) return (long)n.Value;
                var l = raw as long?;
                if (l.HasValue) return l.Value;
                var d = raw as double?;
                if (d.HasValue) return (long)d.Value;
                if (s != null)
                {
                    long parsed;
                    if (long.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                    double dd;
                    if (double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out dd))
                        return (long)dd;
                }
                throw new FormatException("无法解析整数: " + s);
            }

            if (type == "decimal" || type == "double" || type == "float")
            {
                var d = raw as double?;
                if (d.HasValue) return Convert.ToDecimal(d.Value);
                var dec = raw as decimal?;
                if (dec.HasValue) return dec.Value;
                if (s != null)
                {
                    decimal parsed;
                    if (decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析数字: " + s);
            }

            if (col.MaxLength > 0 && s != null && s.Length > col.MaxLength)
                throw new FormatException("超出字段长度 " + col.MaxLength + ": " + s);

            return s ?? raw.ToString();
        }
    }
}
```

- [ ] **Step 4**: `dotnet test ESQLNew.sln` 全绿
- [ ] **Step 5**: 提交 `feat: column mapping + value conversion`

---

### Task 5: 批量写入引擎(BatchInserter)

**Files:**
- Create: `src/ESQLNew/Import/BatchInserter.cs`
- Create: `tests/ESQLNew.Tests/BatchInserterTests.cs`

**Interfaces:**
- Produces:
  - `static string BatchInserter.BuildInsertSql(string table, IList<string> cols, int rowCount)` → `` INSERT INTO `t` (`a`,`b`) VALUES (@p0,@p1),(@p2,@p3) ``
  - `static IEnumerable<object[]> BatchInserter.Split(IEnumerable<object[]> rows, int batchSize)`
  - `static void BatchInserter.Execute(string connStr, string table, IList<ColumnMapping> cols, IEnumerable<object[]> rows, int batchSize, int commitEvery, Action<ImportProgress> onProgress, ImportResult result, System.Threading.CancellationToken ct)`

- [ ] **Step 1**: 写失败测试(纯函数部分)

`tests/ESQLNew.Tests/BatchInserterTests.cs`:
```csharp
using System.Collections.Generic;
using System.Linq;
using ESQLNew.Import;
using Xunit;

namespace ESQLNew.Tests
{
    public class BatchInserterTests
    {
        [Fact]
        public void BuildInsertSql_PlaceholdersCount()
        {
            var sql = BatchInserter.BuildInsertSql("t", new List<string> { "a", "b" }, 3);
            Assert.Equal("INSERT INTO `t` (`a`,`b`) VALUES (@p0,@p1),(@p2,@p3),(@p4,@p5)", sql);
        }

        [Fact]
        public void BuildInsertSql_BacktickEscape()
        {
            var sql = BatchInserter.BuildInsertSql("we`ird", new List<string> { "a" }, 1);
            Assert.StartsWith("INSERT INTO `we``ird`", sql);
        }

        [Fact]
        public void Split_9RowsBatch4_Yields3()
        {
            var rows = Enumerable.Range(0, 9).Select(i => new object[] { i }).ToList();
            var batches = BatchInserter.Split(rows, 4).ToList();
            Assert.Equal(3, batches.Count);
            Assert.Equal(4, batches[0].Length);
            Assert.Equal(1, batches[2].Length);
        }

        [Fact]
        public void Split_Empty_YieldsNone()
        {
            Assert.Empty(BatchInserter.Split(new List<object[]>(0), 4));
        }
    }
}
```

- [ ] **Step 2**: 运行确认失败
- [ ] **Step 3**: 实现纯函数部分

`src/ESQLNew/Import/BatchInserter.cs`(前两方法):
```csharp
using System;
using System.Collections.Generic;
using System.Text;

namespace ESQLNew.Import
{
    public static class BatchInserter
    {
        public static string BuildInsertSql(string table, IList<string> cols, int rowCount)
        {
            var colNames = new List<string>();
            foreach (var c in cols)
                colNames.Add("`" + c.Replace("`", "``") + "`");

            var sb = new StringBuilder();
            sb.Append("INSERT INTO `").Append(table.Replace("`", "``")).Append("` (");
            sb.Append(string.Join(",", colNames));
            sb.Append(") VALUES ");
            int p = 0;
            for (int r = 0; r < rowCount; r++)
            {
                if (r > 0) sb.Append(',');
                sb.Append('(');
                var parts = new string[cols.Count];
                for (int i = 0; i < cols.Count; i++)
                    parts[i] = "@p" + (p++);
                sb.Append(string.Join(",", parts));
                sb.Append(')');
            }
            return sb.ToString();
        }

        public static IEnumerable<object[]> Split(IEnumerable<object[]> rows, int batchSize)
        {
            var buffer = new List<object[]>(batchSize);
            foreach (var row in rows)
            {
                buffer.Add(row);
                if (buffer.Count == batchSize)
                {
                    yield return buffer.ToArray();
                    buffer.Clear();
                }
            }
            if (buffer.Count > 0)
                yield return buffer.ToArray();
        }
    }
}
```

- [ ] **Step 4**: 运行测试确认纯函数通过
- [ ] **Step 5**: 补充 Execute 集成实现(连真实库,事务 + 多值 INSERT + 失败批降级)

```csharp
        public static void Execute(string connStr, string table,
            IList<ColumnMapping> cols, IEnumerable<object[]> rows,
            int batchSize, int commitEvery, Action<ImportProgress> onProgress,
            ImportResult result, System.Threading.CancellationToken ct)
        {
            var matched = new List<ColumnMapping>();
            foreach (var c in cols)
                if (c.Matched) matched.Add(c);

            var fieldNames = new List<string>();
            foreach (var c in matched) fieldNames.Add(c.TableField);

            var sw = System.Diagnostics.Stopwatch.StartNew();
            long total = 0, ok = 0, fail = 0;
            long sinceCommit = 0;
            var conn = new MySql.Data.MySqlClient.MySqlConnection(connStr);
            conn.Open();
            var tx = conn.BeginTransaction();
            try
            {
                foreach (var batch in Split(rows, batchSize))
                {
                    ct.ThrowIfCancellationRequested();
                    int batchOk = 0, batchFail = 0;
                    if (!TryInsertBatch(conn, tx, table, fieldNames, matched, batch, ref batchOk, ref batchFail, result))
                    {
                        foreach (var row in batch)
                        {
                            if (!TryInsertRow(conn, tx, table, fieldNames, matched, row, result))
                                fail++;
                            else
                                ok++;
                        }
                    }
                    else
                    {
                        ok += batchOk;
                        fail += batchFail;
                    }
                    total += batch.Length;
                    sinceCommit += batch.Length;
                    result.BatchCount++;
                    if (sinceCommit >= commitEvery)
                    {
                        tx.Commit();
                        tx.Dispose();
                        tx = conn.BeginTransaction();
                        sinceCommit = 0;
                    }
                    if (onProgress != null)
                        onProgress(new ImportProgress
                        {
                            Processed = total,
                            Succeeded = ok,
                            Failed = fail
                        });
                }
                tx.Commit();
            }
            finally
            {
                tx.Dispose();
                conn.Close();
            }
            sw.Stop();
            result.Total = total;
            result.Succeeded = ok;
            result.Failed = fail;
            result.Elapsed = sw.Elapsed;
            result.RowsPerSecond = sw.Elapsed.TotalSeconds > 0 ? total / sw.Elapsed.TotalSeconds : 0;
        }
```

并补充私有辅助方法 `TryInsertBatch`(拼 BuildInsertSql,参数化后 ExecuteNonQuery,成功返回 true;失败记录整批异常后返回 false)与 `TryInsertRow`(单行参数化插入,失败构造 RowFailure{RowNumber=result.Total+1, Message=ex.Message})。转换用 `ColumnMapper.ConvertValue`,转换异常按失败行处理。

- [ ] **Step 6**: 编译通过(`dotnet build ESQLNew.sln`);Execute 的连库行为留到 Task 6 集成验证
- [ ] **Step 7**: 提交 `feat: batch inserter`

---

### Task 6: 导入编排(ImportEngine)

**Files:**
- Create: `src/ESQLNew/Import/ImportEngine.cs`

**Interfaces:**
- Consumes: ExcelStreamReader、ColumnMapper、BatchInserter、MySqlConnectionBuilder、DbSchemaReader
- Produces:
  - `static IList<ColumnInfo> ImportEngine.GetTableColumns(string connStr, string table)`(SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, CHARACTER_MAXIMUM_LENGTH FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME=@t)
  - `static Task<ImportResult> ImportEngine.Run(string connStr, string table, string excelPath, int batchSize, int commitEvery, Action<ImportProgress> onProgress, System.Threading.CancellationToken ct)`

- [ ] **Step 1**: 实现 `GetTableColumns`(information_schema 查询)
- [ ] **Step 2**: 实现 `Run`:读表头 → GetTableColumns → ColumnMapper.Map → 逐行 ReadRows → 构造参数行 → BatchInserter.Execute;Total 进度用首遍快速计数(如文件大,进度条用"已处理"驱动,MaxValue 自动增长)
- [ ] **Step 3**: 编译 + 单元层面冒烟(依赖真实库的部分本任务仅编译通过,集成在 Task 8 手工验证)
- [ ] **Step 4**: 提交 `feat: import orchestration`

---

### Task 7: 连接 Tab UI

**Files:**
- Modify: `src/ESQLNew/MainForm.cs`(转三 Tab 主窗体)
- Create: `src/ESQLNew/MainForm.Tabs.cs`(partial 类,Tab 构建)

- [ ] **Step 1**: 将 MainForm 改为 partial,构造器初始化 TabControl + 三个 TabPage(连接/导入/日志)
- [ ] **Step 2**: 连接 Tab:字段区(服务器/端口/用户/密码/库名)、RadioButton 切换"表单/完整连接串"、测试连接按钮、保存按钮、高级项(批大小/提交频率/保留日志)
- [ ] **Step 3**: 测试连接:Task.Run 建连,成功 MessageBox(标题含绿 ✅ U+2705)失败(红 ❌ U+274C)+ 异常信息
- [ ] **Step 4**: 启动从 `%APPDATA%\ESQLNew\config.json` 载入;保存写入;提供连接串生成到文本框供查看
- [ ] **Step 5**: `dotnet build ESQLNew.sln` 通过;提交 `feat: connection tab`

---

### Task 8: 导入 Tab UI(文件/表名/预览/进度/报告)

**Files:**
- Modify: `src/ESQLNew/MainForm.Tabs.cs`

- [ ] **Step 1**: 选文件 OpenFileDialog(过滤 `Excel 文件|*.xlsx;*.xls`);选定后 `ExcelStreamReader.SheetNames` 首个 sheet 名预填表名文本框
- [ ] **Step 2**: 匹配预览:读表字段 + 表头 → 映射表(DataGridView 显示 Excel 列/表字段/匹配状态)+ 样例区(前 100 行,只显示匹配列)
- [ ] **Step 3**: 开始导入:后台 Task.Run(ImportEngine.Run)+ `BeginInvoke` 刷新 ProgressBar(用已处理行数驱动,MaxValue 随已读数增长)/计数标签
- [ ] **Step 4**: 完成后弹汇总报告对话框:总数/成功/失败/耗时/行每秒/批次数 + "导出失败明细 CSV"按钮
- [ ] **Step 5**: 校验表名 `^[A-Za-z0-9_]+$`,非法则提示
- [ ] **Step 6**: 提交 `feat: import tab`

---

### Task 9: 日志 Tab(SqLiteLogStore)

**Files:**
- Create: `src/ESQLNew/Logging/SqLiteLogStore.cs`
- Modify: `src/ESQLNew/MainForm.Tabs.cs`

**Interfaces:**
- Produces:
  - `SqLiteLogStore { string DbPath; void LogImport(DateTime ts, string file, string table, long total, long ok, long fail, long ms, double rps, System.Collections.Generic.List<RowFailure>); System.Collections.Generic.List<object[]> Query(int page, int pageSize); void PurgeOld(int days); void EnsureSchema(); }`
  - DbPath 固定 `%APPDATA%\ESQLNew\logs.db`

- [ ] **Step 1**: EnsureSchema 建表 `LogImport(Id INTEGER PRIMARY KEY AUTOINCREMENT, Timestamp TEXT, FileName TEXT, TableName TEXT, Total INTEGER, Succeeded INTEGER, Failed INTEGER, ElapsedMs INTEGER, RowsPerSecond REAL)` 与 `LogFailure(Id INTEGER PRIMARY KEY AUTOINCREMENT, ImportId INTEGER, RowNumber INTEGER, Message TEXT)`
- [ ] **Step 2**: 实现 LogImport / Query / PurgeOld(删除超过 LogRetentionDays 的 LogImport 及其 LogFailure)
- [ ] **Step 3**: 日志 Tab:开关(勾选=保留,绑定 KeepLogs)+ DataGridView 倒序列表 + 选中行显示失败明细
- [ ] **Step 4**: 导入完成后若 KeepLogs 则写日志
- [ ] **Step 5**: 提交 `feat: sqlite log store`

---

### Task 10: 性能实测与调优

**Files:**
- Create: `docs/perf-report.md`

- [ ] **Step 1**: 用用户提供的全量 Excel 文件,分别测 1万/10万/百万行
- [ ] **Step 2**: 对比批大小 500/2000/5000、提交频率 2000/5000/20000,记录行/秒(汇总报告自动给出)
- [ ] **Step 3**: 记录结果到 `docs/perf-report.md`(数据量/批大小/耗时/行每秒/失败数)
- [ ] **Step 4**: 若实测发现默认值非最优,调整 `AppConfig` 默认值并提交

---

### Task 11: 验收与收尾

- [ ] **Step 1**: 全量构建 `dotnet build ESQLNew.sln`(Release),确认输出 exe + 依赖 dll 齐备
- [ ] **Step 2**: 人工走查:连接测试✅/❌、预览匹配、导入进度、失败明细、日志页、配置持久化
- [ ] **Step 3**: 确认 xcopy 部署目录可直接运行(x64)
- [ ] **Step 4**: 提交最终状态;整理交付说明

---

## 模板文件已知问题

`历史维修安装数据.xlsx` 的 sheet1 deflate 流读取时于 491520 字节处中断(疑似保存不完整/被占用)。Task 3 会先以该文件验证;若 ExcelDataReader 同样中断,则标记模板损坏,性能实测改用用户提供的全量文件。
