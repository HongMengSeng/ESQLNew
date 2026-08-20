# 可配置目录存储配置文件实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 config.json / column-maps.json / logs.db 三个配置文件从 `%APPDATA%\ESQLNew` 迁移到可配置目录(默认 exe 所在目录,可由 exe 旁 app-path.txt 指定),并首次启动自动从旧位置迁移。

**Architecture:** 新增静态 `AppPaths` 类统一计算配置目录与各文件路径(默认 exe 目录,app-path.txt 可覆盖);三处硬编码路径改为引用 `AppPaths`;启动时做一次旧→新迁移复制。

**Tech Stack:** .NET Framework 4.6.2 WinForms, xUnit。

## Global Constraints

- net462, LangVersion 7.3
- 代码不写任何注释
- UI 文案为中文
- 现有 56 个测试保持通过
- 每个 Task 结束跑 `dotnet test ESQLNew.sln`,预期全绿
- 运行中 ESQLNew.exe 会锁 exe 导致 MSB3027/MSB3021 → 先 `Stop-Process -Name ESQLNew -Force`
- 仅修改本计划列出的文件

---

### Task 1: AppPaths 类 + 单元测试

**Files:**
- Create: `src/ESQLNew/Core/AppPaths.cs`
- Create: `tests/ESQLNew.Tests/AppPathsTests.cs`

**Interfaces:**
- Produces:
  - `public static class AppPaths`
  - `public static string ConfigDir { get; }`
  - `public static string ConfigFile { get; }`
  - `public static string ColumnMapsFile { get; }`
  - `public static string LogDbFile { get; }`
  - `public static string LegacyDir { get; }`
  - `public static void Initialize(string baseDirectory, string appPathTxtPath, string legacyDir)` — 显式初始化(可测试);若未调用则用默认(exe 目录 + 无 app-path.txt + %APPDATA%\ESQLNew)
  - `public static void MigrateFromLegacy()` — 把 LegacyDir 中缺少于 ConfigDir 的文件复制过去

**设计要点(可测试性):**
- `AppPaths` 用静态字段存储解析结果,`Initialize(baseDirectory, appPathTxtPath, legacyDir)` 在测试中可注入临时目录,避免依赖真实 exe 路径。
- 默认(未调用 Initialize):`baseDirectory = AppDomain.CurrentDomain.BaseDirectory`,`appPathTxtPath = baseDirectory\app-path.txt`,`legacyDir = %APPDATA%\ESQLNew`。
- `ConfigDir` 解析:若 `appPathTxtPath` 存在且首行非空 → 规范化该路径;否则 baseDirectory。若目录不存在 → `Directory.CreateDirectory`(失败则回退 baseDirectory)。
- `MigrateFromLegacy`:对 `config.json`、`column-maps.json`、`logs.db` 三个文件,若 `ConfigDir` 下不存在而 `LegacyDir` 下存在,则 `File.Copy` 到 ConfigDir。幂等。

- [ ] **Step 1: 写失败测试**

创建 `tests/ESQLNew.Tests/AppPathsTests.cs`:

```csharp
using System;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class AppPathsTests : IDisposable
    {
        private readonly string _tmp;
        public AppPathsTests()
        {
            _tmp = Path.Combine(Path.GetTempPath(), "apppaths_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tmp);
        }
        public void Dispose()
        {
            try { Directory.Delete(_tmp, true); } catch { }
        }

        [Fact]
        public void ConfigDir_DefaultsToBaseDirectory()
        {
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(_tmp, AppPaths.ConfigDir);
        }

        [Fact]
        public void ConfigDir_UsesAppPathTxtWhenPresent()
        {
            var custom = Path.Combine(_tmp, "custom");
            File.WriteAllText(Path.Combine(_tmp, "app-path.txt"), custom);
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.GetFullPath(custom), Path.GetFullPath(AppPaths.ConfigDir));
            Assert.True(Directory.Exists(AppPaths.ConfigDir));
        }

        [Fact]
        public void ConfigDir_AppPathTxtEmpty_FallsBackToBase()
        {
            File.WriteAllText(Path.Combine(_tmp, "app-path.txt"), "   ");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.GetFullPath(_tmp), Path.GetFullPath(AppPaths.ConfigDir));
        }

        [Fact]
        public void FileProperties_PointIntoConfigDir()
        {
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.Combine(_tmp, "config.json"), AppPaths.ConfigFile);
            Assert.Equal(Path.Combine(_tmp, "column-maps.json"), AppPaths.ColumnMapsFile);
            Assert.Equal(Path.Combine(_tmp, "logs.db"), AppPaths.LogDbFile);
        }

        [Fact]
        public void MigrateFromLegacy_CopiesMissingFiles()
        {
            var legacy = Path.Combine(_tmp, "legacy");
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "config.json"), "cfg");
            File.WriteAllText(Path.Combine(legacy, "logs.db"), "log");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), legacy);
            AppPaths.MigrateFromLegacy();
            Assert.True(File.Exists(Path.Combine(_tmp, "config.json")));
            Assert.True(File.Exists(Path.Combine(_tmp, "logs.db")));
            Assert.Equal("cfg", File.ReadAllText(Path.Combine(_tmp, "config.json")));
        }

        [Fact]
        public void MigrateFromLegacy_DoesNotOverwriteExisting()
        {
            var legacy = Path.Combine(_tmp, "legacy");
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "config.json"), "legacy");
            File.WriteAllText(Path.Combine(_tmp, "config.json"), "new");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), legacy);
            AppPaths.MigrateFromLegacy();
            Assert.Equal("new", File.ReadAllText(Path.Combine(_tmp, "config.json")));
        }
    }
}
```

- [ ] **Step 2: 运行测试确认失败**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~AppPaths"`
Expected: 编译失败(AppPaths 不存在)→ 红。

- [ ] **Step 3: 实现 AppPaths**

创建 `src/ESQLNew/Core/AppPaths.cs`:

```csharp
using System;
using System.IO;

namespace ESQLNew.Core
{
    public static class AppPaths
    {
        private static string _configDir;
        private static string _configFile;
        private static string _columnMapsFile;
        private static string _logDbFile;
        private static string _legacyDir;

        public static string ConfigDir { get { return _configDir; } }
        public static string ConfigFile { get { return _configFile; } }
        public static string ColumnMapsFile { get { return _columnMapsFile; } }
        public static string LogDbFile { get { return _logDbFile; } }
        public static string LegacyDir { get { return _legacyDir; } }

        public static void Initialize(string baseDirectory, string appPathTxtPath, string legacyDir)
        {
            string dir = baseDirectory;
            try
            {
                if (File.Exists(appPathTxtPath))
                {
                    string line = null;
                    using (var r = new StreamReader(appPathTxtPath))
                        line = r.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                        dir = Path.GetFullPath(line.Trim());
                }
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
            catch
            {
                dir = baseDirectory;
            }
            _configDir = dir;
            _configFile = Path.Combine(dir, "config.json");
            _columnMapsFile = Path.Combine(dir, "column-maps.json");
            _logDbFile = Path.Combine(dir, "logs.db");
            _legacyDir = legacyDir;
        }

        public static void MigrateFromLegacy()
        {
            try
            {
                if (string.IsNullOrEmpty(_legacyDir) || !Directory.Exists(_legacyDir)) return;
                string[] files = { "config.json", "column-maps.json", "logs.db" };
                foreach (var f in files)
                {
                    string src = Path.Combine(_legacyDir, f);
                    string dst = Path.Combine(_configDir, f);
                    if (File.Exists(src) && !File.Exists(dst))
                        File.Copy(src, dst);
                }
            }
            catch
            {
            }
        }

        static AppPaths()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            Initialize(baseDir, Path.Combine(baseDir, "app-path.txt"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ESQLNew"));
        }
    }
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `dotnet test ESQLNew.sln --filter "FullyQualifiedName~AppPaths"`
Expected: PASS。

- [ ] **Step 5: 运行全部测试 + 提交**

Run: `dotnet test ESQLNew.sln` → 全绿(原 56 + 新增 6 = 62)。
Commit:
```bash
git add src/ESQLNew/Core/AppPaths.cs tests/ESQLNew.Tests/AppPathsTests.cs
git commit -m "feat: add AppPaths configurable config directory"
```

---

### Task 2: 接入三处路径 + 启动迁移

**Files:**
- Modify: `src/ESQLNew/Core/ColumnMapStore.cs`
- Modify: `src/ESQLNew/MainForm.Tabs.cs`
- Modify: `src/ESQLNew/Program.cs`(启动时调用迁移)

**Interfaces:**
- Consumes: `AppPaths.ConfigFile`、`AppPaths.ColumnMapsFile`、`AppPaths.LogDbFile`、`AppPaths.MigrateFromLegacy()`(Task 1)
- Produces: 无新公开接口

#### 改造点

1. **ColumnMapStore.ConfigPath** → `return AppPaths.ColumnMapsFile;`
   - 在 `src/ESQLNew/Core/ColumnMapStore.cs` 中,把 ConfigPath getter 内的 `Path.Combine(ApplicationData, "ESQLNew", "column-maps.json")` 替换为 `return AppPaths.ColumnMapsFile;`。删除不再使用的 `using System.IO`(若不再需要)。

2. **MainForm.Tabs.cs `ConfigPath`**(config.json)→ `return AppPaths.ConfigFile;`

3. **MainForm.Tabs.cs `LogStore`** → 用 `AppPaths.LogDbFile`:
   ```csharp
   _logStore = new SqLiteLogStore(AppPaths.LogDbFile);
   ```

4. **Program.cs 启动迁移**:在 Main 开头(或 MainForm 构造前)调用 `AppPaths.MigrateFromLegacy();`。

#### 测试

本 Task 是路径替换 + 启动调用,不新增单元测试(路径行为已由 Task 1 测试覆盖)。验证方式:构建 0 错误 + 现有测试全绿 + 手动确认配置文件生成在 exe 目录。

- [ ] **Step 1: 改造三处路径**(上述 1-3)

- [ ] **Step 2: Program.cs 调用迁移**(上述 4)

- [ ] **Step 3: 构建 + 测试**

Run: `dotnet build ESQLNew.sln -c Release`(0 错误)→ `dotnet test ESQLNew.sln`(全绿)

- [ ] **Step 4: 真实验证**

启动发布版,确认:
- exe 目录下生成 `config.json`、`column-maps.json`(以及首次启动若旧 %APPDATA%\ESQLNew 有数据则自动迁移)
- `%APPDATA%\ESQLNew` 的旧配置被复制到 exe 目录
- 用 probe 或启动应用验证配置读写正常

- [ ] **Step 5: 提交**

```bash
git add src/ESQLNew/Core/ColumnMapStore.cs src/ESQLNew/MainForm.Tabs.cs src/ESQLNew/Program.cs
git commit -m "feat: route config files through AppPaths with legacy migration"
```
