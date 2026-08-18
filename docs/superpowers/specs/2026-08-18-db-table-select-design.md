# ESQLNew 数据库/表选择功能设计

日期:2026-08-18
状态:待审阅

## 背景

当前连接 Tab 的"数据库"是手动输入文本框,导入 Tab 的"目标表"也是手动输入。用户希望连接成功后能从列表选择数据库、并拉出该库的所有表,直接用于导入。

## 需求(经逐项澄清确认)

1. **选库+选表都要**:连接 Tab 选库,导入 Tab 选表,预览/导入直接使用选中的库和表。
2. **拉取时机**:测试连接成功后自动拉取库列表;切换库时自动重拉表列表;同时提供手动刷新按钮(自动+手动都要)。
3. **选表形式**:导入 Tab 目标表改为可搜索下拉列表(可输入可下拉),选库后自动填充;导入时校验所选表名仍存在。
4. **库列表范围**:只列当前用户有权限的库,可搜索;默认隐藏系统库(mysql/information_schema/performance_schema/sys),提供"显示系统库"开关可切回。
5. **表列表联动**:连接 Tab 选库即拉表,导入 Tab 的目标表下拉直接用已拉好的列表(切换库自动重拉)。
6. **无库连接**:允许连接串不含数据库名也能测试连接;成功后再从库列表选库并重拉表列表。

## 方案A:可搜索下拉联动选库选表(已选定)

### 架构与组件

新增一个元数据查询类,负责向 MySQL 拉取库/表清单,保持与现有 `ImportEngine.GetTableColumns` 同一风格。

**新增 `src/ESQLNew/Core/DbMetadata.cs`(静态类):**

- `static IList<string> GetDatabases(string connStr)` — 执行 `SHOW DATABASES`,返回当前用户可访问的所有库名。
- `static IList<string> GetTables(string connStr, string database)` — 执行 `SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME`,参数化查询(防注入),只返回基表。
- `static bool IsSystemDatabase(string name)` — 判断是否系统库(mysql/information_schema/performance_schema/sys,大小写不敏感)。
- 两个列表方法均在 `try/catch` 内将底层异常包装为带中文提示的 `InvalidOperationException`(与 ImportEngine 风格一致)。

**修改 `src/ESQLNew/Core/MySqlConnectionBuilder.cs`:**

- `Build(...)` 中 database 参数为空或空白时,连接串省略 `Database=...` 段(允许无库连接)。
- 其余连接串参数不变。

**修改 `src/ESQLNew/Core/AppConfig.cs`:**

- 新增属性 `bool ShowSystemDatabases { get; set; } = false;`(默认隐藏系统库)。
- 其余字段与序列化逻辑不变。

### 连接 Tab 改造

- "数据库"输入框 `_databaseTextBox` 改为可搜索下拉 `_databaseComboBox`(ComboBox,`DropDownStyle = DropDown` 即可输入可下拉,`AutoCompleteSource = ListItems`, `AutoCompleteMode = SuggestAppend`)。
- **测试连接成功后**:
  1. 用当前连接串(可能无数据库)调用 `DbMetadata.GetDatabases` 拉取库列表;
  2. 按 `AppConfig.ShowSystemDatabases` 过滤系统库;
  3. 填充下拉框;
  4. 若原 `_databaseTextBox` 中的文本匹配某个库,则选中该项;否则文本保留为空(不自动塞入)。
- 切换库下拉(`SelectedIndexChanged`,即用户从下拉选中一项时):若连接成功,自动调用 `DbMetadata.GetTables(connStr, 库名)` 拉取该库的表列表,并写入导入 Tab 的目标表下拉。手输库名不触发自动拉表(用户需点"刷新表"按钮),避免每次按键都发查询。
- 新增"刷新"按钮(库列表)与"显示系统库"复选框:切换复选框时立即重拉库列表(过滤逻辑按新开关执行)。
- 新增"刷新表"按钮(导入 Tab):重新拉取当前所选库的表列表。

### 导入 Tab 改造

- "目标表"输入框 `_tableTextBox` 改为可搜索下拉 `_tableComboBox`(ComboBox,`DropDown` 风格,可输入可下拉)。
- 内容来源:连接 Tab 当前选中库的 `GetTables` 结果;连接 Tab 切库时自动填充。
- 预览/导入流程:
  - 表名从下拉文本取值,仍沿用现有 `^[A-Za-z0-9_]+$` 校验;
  - 预览前(GetTableColumns 前)额外校验:若下拉非空且当前库已拉取过表列表,校验所选表名在列表中存在,不存在则提示"表名不存在或已更改,请重新选择";
  - 导入前同样校验。
- 未拉取过表列表时(未连接/未选库),下拉为空,预览/导入时按现有逻辑提示"请先选择 Excel 文件/目标表"。

### 数据流

```
用户输入服务器/端口/账号/密码 → 测试连接
  → 连接成功 → DbMetadata.GetDatabases → 填充库下拉(按开关过滤系统库)
用户选择库 → 自动 DbMetadata.GetTables(库) → 填充导入 Tab 表下拉
用户点刷新 → 重拉库/表列表
用户选表/输入表名 → 预览(GetTableColumns) / 导入(ImportEngine.Run)
```

### 错误处理

- 测试连接失败:不拉列表,库/表下拉保持原样(不清空),错误提示不变。
- 拉库失败:库下拉不动,状态栏/MessageBox 提示"无法读取数据库列表:<原因>"。
- 拉表失败:表下拉清空并提示"无法读取表列表:<原因>"(库可能已删或权限变化)。
- 所有列表拉取在 UI 线程执行(与预览一致),量大时仅在刷新动作期间禁用对应按钮。

### 测试

- `MySqlConnectionBuilderTests`:新增用例——database 为空/空白时连接串不含 `Database=`;非空时包含。
- `DbMetadata` 的纯逻辑(`IsSystemDatabase`)新增单元测试:大小写不敏感、四个系统库命中、普通库不命中。
- 列表拉取依赖真实 MySQL,不做单元测试(与 `GetTableColumns` 一致);由手动验收覆盖。
- 既有 24 个测试保持全绿。

### 边界与不做的事

- 不做数据库/表的搜索框 UI 组件,仅用 ComboBox 自带 AutoComplete。
- 不做"显示表注释/列信息"等扩展(现有预览已覆盖列匹配)。
- 不做多选/批量导入。
- 不做记住"上次选中表名"的持久化(超出本次需求)。
- 表列表仅列 BASE TABLE(不含视图),与导入目标一致。
- 下拉框手输库名不自动拉表(避免逐键查询);库名/表名下拉选项取自 `SelectedIndexChanged` 事件。

## 验收标准

1. 连接串无数据库名时测试连接成功,库下拉出现该账号可见的库。
2. 默认不显示系统库;勾选"显示系统库"后出现 mysql 等系统库。
3. 选择某库后,导入 Tab 目标表下拉自动出现该库所有基表。
4. 手动刷新按钮可重新拉取库/表列表。
5. 表名可下拉选择也可手动输入;预览/导入校验所选表名存在且合法。
6. 既有功能(连接测试、预览、导入、日志)不回归。