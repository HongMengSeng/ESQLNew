# 可配置目录存储配置文件设计

日期:2026-08-19
状态:已批准

## 背景与问题

程序当前把三个配置文件硬编码到 `%APPDATA%\ESQLNew`:
- `config.json`(数据库连接)
- `column-maps.json`(列映射)
- `logs.db`(SQLite 导入日志)

用户要把程序分发给别人使用。若配置留在 %APPDATA%,对方换机器配置丢失、也难以找到/修改/随程序分发。目标:配置文件移到**可配置目录**,默认程序(exe)所在目录,使分发时整个文件夹即开即用。

## 需求

1. 配置文件存放到一个**可配置目录**。
2. 默认用 **exe 所在目录**(分发即用)。
3. 可通过 **exe 同目录的小配置文件**(`app-path.txt`)指定其它配置目录;不存在则该文件 → 默认 exe 目录。
4. 三个文件(`config.json`、`column-maps.json`、`logs.db`)全部移入该目录。
5. 兼容旧安装:若新位置无配置而 `%APPDATA%\ESQLNew` 有旧配置,首次启动自动复制过去。

## 架构

### 组件

**新增 `src/ESQLNew/Core/AppPaths.cs`**(静态类,统一路径来源):
- `public static string ConfigDir { get; }` — 配置目录
- `public static string ConfigFile { get; }` — `ConfigDir\config.json`
- `public static string ColumnMapsFile { get; }` — `ConfigDir\column-maps.json`
- `public static string LogDbFile { get; }` — `ConfigDir\logs.db`
- `public static string LegacyDir { get; }` — 旧 `%APPDATA%\ESQLNew`(用于迁移)

**ConfigDir 解析逻辑**:
1. exe 同目录存在 `app-path.txt` → 读取首行路径;非空则规范化,作为 ConfigDir。
2. 否则默认 `AppDomain.CurrentDomain.BaseDirectory`(exe 所在目录)。
3. 若目录不存在,尝试 `Directory.CreateDirectory`。

### 改造点

1. **`ColumnMapStore.ConfigPath`** → 返回 `AppPaths.ColumnMapsFile`。
2. **`MainForm.Tabs.cs` 的 `ConfigPath`**(config.json)→ 返回 `AppPaths.ConfigFile`。
3. **`MainForm.Tabs.cs` 的 `LogStore`**(logs.db)→ 用 `AppPaths.LogDbFile`。
4. **首次迁移**:程序启动时(或首次访问路径时)检查:若 `ConfigDir` 下缺少某文件且 `LegacyDir` 下存在同名文件,则复制到 ConfigDir。只做一次,幂等。

## 数据流

```
启动 → AppPaths 解析 ConfigDir(exe旁app-path.txt 或 默认exe目录)
  → 首次:若 ConfigDir 无某文件而 LegacyDir 有 → 自动复制
  → config.json / column-maps.json / logs.db 均从 ConfigDir 读写
```

## 错误处理

- `app-path.txt` 内容为空/无效 → 回退默认 exe 目录。
- 配置目录无法创建 → 记录并回退 exe 目录(尽力而为)。
- 旧配置复制失败 → 不影响启动,跳过(不阻塞)。

## 测试

- `AppPaths.ConfigDir` 默认 = exe 目录。
- 存在 `app-path.txt` 时 = 文件指定路径。
- 迁移:LegacyDir 有文件而 ConfigDir 无 → 复制后存在;已有 → 不覆盖。
- 现有 56 个测试保持通过。

## 兼容性

- 新部署:配置直接在 exe 目录,即开即用。
- 旧部署:首次启动自动从 %APPDATA% 复制已有配置到新目录。
- `app-path.txt` 可选;不提供则默认 exe 目录。
