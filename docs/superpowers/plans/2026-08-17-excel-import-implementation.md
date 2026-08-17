# Excel 批量导入 MySQL 工具 —— 实现计划

- 日期:2026-08-17
- 关联设计:`docs/superpowers/specs/2026-08-17-excel-import-design.md`
- 目标框架:.NET Framework 4.6.2 + WinForms + C#
- 状态:待执行

## 0. 实现顺序总览

按"先底层后 UI、先核心链路后辅助功能"的顺序推进,每阶段产出可独立编译运行/验证的产物。共 7 个阶段,每阶段标注预计文件与验收点。

| 阶段 | 主题 | 产出 | 验收 |
|------|------|------|------|
| 1 | 工程骨架与依赖 | csproj + packages.config + App.config | 能编译空 WinForms 窗体并启动 |
| 2 | 配置与连接层 | AppConfig、ConnectionService | 测试连接按钮工作,配置持久化 |
| 3 | Excel 读取与列匹配 | ExcelStreamReader、ColumnMapper | 能读列头与前 100 行,匹配关系可见 |
| 4 | 批量写入与失败降级 | BatchInserter | 单元覆盖批拆分、多值 INSERT、批失败逐行定位 |
| 5 | 导入编排与进度 | ImportService + Tab2 导入 UI | 端到端导入小文件,进度与汇总正确 |
| 6 | 本地日志与 Tab3 | LogService(SQLite)+ 日志 UI | 导入后日志可见,失败明细可展开,归档生效 |
| 7 | 调优与验收 | 性能实测与参数文档 | 1万/10万/百万级行实测达标 |

---

## 阶段 1:工程骨架与依赖

**目标**:搭建可编译运行的 WinForms 工程,补齐所有外部依赖。

**产出文件**
- `ExcelImportTool.sln`
- `src/ExcelImportTool/ExcelImportTool.csproj`
  - TargetFrameworkVersion = v4.6.2
  - OutputType = WinExe
  - 引用:WinForms / System.Data / System.Configuration
- `src/ExcelImportTool/packages.config`(以下 NuGet 包)
  - Microsoft.NETFramework.ReferenceAssemblies.net462
  - MySql.Data (8.0.x)
  - ExcelDataReader (3.6.x) + ExcelDataReader.DataSet
  - System.Data.SQLite (含 x86/x64 两套原生 dll)
- `src/ExcelImportTool/App.config`
  - `<startup>` supportedRuntime 钉到 4.5/4.6
  - `<runtime>` `<gcServer enabled="true">`(服务器 GC 利于大批量内存)
  - SQLite `<system.data>` DbProviderFactory 注册
  - 连接串占位
- `src/ExcelImportTool/Program.cs`(Main,启用文字渲染、异常全局兜底)
- `src/ExcelImportTool/Forms/MainForm.cs` + `.Designer.cs`
  - 3 个 Tab 的空壳:`tabConnection` / `tabImport` / `tabLog`
- `src/ExcelImportTool/Properties/AssemblyInfo.cs`、`app.manifest`(兼容性 Windows 10/11)

**验收**
- `msbuild` 能 Clean + Build 通过
- F5 启动看到 3 个 Tab 空白页

**注意**
- System.Data.SQLite 需保证运行时 `x86`/`x64` 子目录的 `SQLite.Interop.dll` 随输出目录复制(在 csproj 用 `<None Include="x86\..."><CopyToOutputDirectory>PreserveNewest</...>` 或通过 NuGet 的 `build/` targets 自动处理,优先依赖 NuGet 自带 targets)
- MySql.Data 8.x 在 net462 下需注意 TLS 配置,必要时在 App.config 显式 `Service Manager` 关闭 SSL 或用 `SslMode=none/Required` 按用户环境后续在 ConnectionService 暴露

---

## 阶段 2:配置与连接层

**目标**:实现本地配置读写 + MySQL 连接管理,完成 Tab1 交互。

**产出文件**
- `src/ExcelImportTool/Config/AppConfig.cs`
  - 静态方法:`Load()` / `Save(AppConfigDto)` / `Get<T>(key, default)` / `Set(key, value)`
  - 存储位置:`%AppData%\ExcelImportTool\config.json`(用 `Environment.GetFolderPath`)
  - JSON 序列化用内置 `System.Web.Script.Serialization` 或自带极简 JSON(避免再引依赖);优先 `System.Web.Script.Serialization.JavaScriptSerializer`(.NET 4.6 自带)
- `src/ExcelImportTool/Config/AppConfigDto.cs`
  - 字段:Host、Port(3306)、User、Password、Database、RawConnectionString(完整串模式)、LastTable、BatchSize(2000)、CommitInterval(5000)、KeepLog(true)、LogRetentionDays(30)
- `src/ExcelImportTool/Services/ConnectionService.cs`
  - `BuildConnectionString(AppConfigDto)`:表单模式拼装;RawConnectionString 模式直接用用户串
  - `TestConnection(string connStr)` 返回 `(bool ok, string message)`
  - `GetCurrent()` / `SaveCurrent(AppConfigDto)`:委托 AppConfig
  - 连接串参数统一处理:Charset=utf8;Pooling=true;默认 `SslMode=Preferred`(允许失败时降级,由 UI 提示开关)
- `src/ExcelImportTool/Forms/MainForm.ConnectionTab.cs`(partial)
  - 模式切换 RadioButton:表单 / 完整串
  - 输入控件 + 「测试连接」「保存」按钮
  - 测试结果 Label,成功绿底 ✅,失败红底 ❌ + 错误信息
  - 启动时 LoadCurrent 回填

**验收**
- 表单模式填入测试库凭据 → 测试连接 ✅
- 完整串模式粘贴串 → 测试连接 ✅
- 保存 → 关闭程序重启 → 字段自动回填
- 故意填错 → 红色 ❌ 并显示 MySQL 异常信息

**注意**
- 不要把密码明文写到工程目录;AppData 路径避免权限问题
- 测试连接放在后台 `Task.Run`,避免 MySQL 默认超时阻塞 UI;UI 用 `Invoke` 回写结果
- MySql.Data 8.0 连接若报 SSL 错误,在 ConnectionService 暴露 `SslMode` 与 `AllowPublicKeyRetrieval` 开关由用户在表单模式里勾选

---

## 阶段 3:Excel 读取与列匹配

**目标**:流式读取 Excel(不整表加载),实现列名↔表字段匹配,完成 Tab2 预览区。

**产出文件**
- `src/ExcelImportTool/Services/Excel/ExcelStreamReader.cs`
  - 构造:`new ExcelStreamReader(filePath)`
  - `IReadOnlyList<string> ReadHeader()`:打开流,定位第一个 sheet,读列头(默认第 1 行)
  - `IEnumerable<object[]> EnumerateRows()`:前向只读 `IExcelDataReader.Read()` 逐行产出;不调用 `AsDataSet()`
  - `long TotalRows`:通过预读一次计数或元信息;百万级时先快扫一遍计数(只读流,耗时可接受),用于进度分母
  - `Dispose()`:释放 `IExcelDataReader` 与底层 FileStream
  - 支持 .xls(FORMAT)与 .xlsx;空 sheet 与首行空列头抛明确异常
- `src/ExcelImportTool/Services/Schema/TableSchemaService.cs`
  - `GetColumns(string connStr, string tableName) -> IReadOnlyList<ColumnInfo>`
    - 用 `information_schema.columns` 查:列名、数据类型、是否可空、字符长度、主键标记
    - 列名比较用 `OrdinalIgnoreCase`
- `src/ExcelImportTool/Services/Mapping/ColumnMapper.cs`
  - `Map(IReadOnlyList<string> excelHeader, IReadOnlyList<ColumnInfo> tableColumns) -> IReadOnlyList<ColumnMap>`
  - 规则:精确匹配 → 忽略大小写 → 去除下划线/空格归一化匹配 → 仍无则 `Matched=false`
  - 返回包含:Excel 列索引、Excel 列名、表列名、是否匹配、目标类型
  - 提供 `BuildInsertSkeleton(columnMaps, tableName)`:产出多值 INSERT 模板字符串与参数名列表,供 BatchInserter 复用
- `src/ExcelImportTool/Forms/MainForm.ImportTab.Preview.cs`(partial)
  - 「选择文件」按钮 + 路径文本框 + 过滤器 `*.xlsx;*.xls`
  - 目标表名输入框(历史记忆最近 10 个,下拉)
  - 「读取表结构 + 预览」按钮:并行取表结构 + 读 Excel 头与前 100 行,展示映射表与样例 DataGridView
  - 未匹配列显示灰底"未匹配,导入时忽略"

**验收**
- 选文件 → 输入表名 → 预览正确展示映射与前 100 行
- 表不存在 → MessageBox 提示并中止
- 列名与表字段大小写/下划线差异 → 仍能自动匹配
- 读取百万行文件时不卡死不爆内存(此处仅读头 + 100 行,要快)

**注意**
- ExcelDataReader 在 .xls 需 `AsDataSet` 才能取行数,百万行会爆;改为"只读流计数",即遍历 `Read()` 累加,或仅在用户点开始导入后再计数。预览阶段只读前 100 行不计数
- 路径含中文/空格需正确处理;文件被占用(Excel 打开中)给出友好提示

---

## 阶段 4:批量写入与失败降级

**目标**:实现核心写入引擎,不依赖 UI,可被单元测试覆盖。

**产出文件**
- `src/ExcelImportTool/Services/Import/BatchInserter.cs`
  - 构造注入:`connStr`、`tableName`、`columnMaps`(匹配成功的列)、`batchSize`、`commitInterval`
  - `InsertBatch(IDbConnection conn, IDbTransaction tx, IReadOnlyList<object[]> rows) -> BatchResult`
    - 构造多值 INSERT:`INSERT INTO t (c1,c2) VALUES (@r0_c1,@r0_c2),(@r1_c1,...)`
    - 参数化,避免注入与类型转换误差
    - `BatchResult`:SuccessCount、FailedRows(List<行号+错误>)
  - 失败降级:`InsertBatch` 抛异常 → 调用 `InsertRowWise` 逐行插入,逐行 try/catch 收集失败明细,成功的仍提交到 tx
  - 事务管理:由上层 ImportService 控制 `BeginTransaction`/`Commit`;BatchInserter 只在传入事务上下文内执行
  - 提供 `PrepareInsertSql(int rowCount)` 缓存 SQL 文本,避免每批拼字符串
- `src/ExcelImportTool/Services/Import/ImportResult.cs`
  - 汇总结构:Total、Success、Failed、Elapsed、RowsPerSecond、BatchCount、AvgBatchMs、FailedRows(List<行号、原始数据摘要、错误>)
- `tests/ExcelImportTool.Tests/Mapping/ColumnMapperTests.cs`
- `tests/ExcelImportTool.Tests/Import/BatchInserterTests.cs`
  - 用 SQLite in-memory 或 Mock IDbConnection 模拟批拆分与失败降级路径,不依赖真实 MySQL

**验收**
- 单元测试通过:批拆分边界(行数=批大小、批大小+1)、空批、全失败批降级、部分失败批
- 多值 INSERT SQL 模板在 MySQL 客户端可直接执行(参数替换后)

**注意**
- MySQL 多值 INSERT 的 `max_allowed_packet` 默认 4MB,2000 行 × N 列可能超限;在 ConnectionService 暴露 `max_allowed_packet` 检测,超限时自动折半批大小重试一次,仍超则报错
- 参数命名带行索引避免冲突:`@r{rowIndex}_c{colIndex}`
- 时间/Decimal 类型由 MySql.Data 参数转换;DBNull 处理为 `DBNull.Value`

---

## 阶段 5:导入编排与进度(Tab2 主链路)

**目标**:打通"读 Excel → 匹配 → 写入 → 汇总"全链路,UI 进度实时反馈。

**产出文件**
- `src/ExcelImportTool/Services/Import/ImportService.cs`
  - 事件:`event EventHandler<ProgressInfo> ProgressChanged`(`Processed`、`Success`、`Failed`、`Total`、`Phase`)
  - `ImportAsync(ImportRequest req, CancellationToken ct) -> ImportResult`
    - 打开 ExcelStreamReader
    - 调 ColumnMapper 生成 columnMaps(若用户在预览里手改过映射,优先用户映射)
    - 调 TableSchemaService 取表结构校验
    - open conn → 循环:
      - 从 ExcelStreamReader 读取一批 batchSize 行
      - 凑到 commitInterval 行触发 `BeginTransaction` → 累计 2~3 批后 `Commit`
      - 调 BatchInserter.InsertBatch;失败降级
      - 上报 ProgressChanged
    - 结束收尾 Commit 剩余事务,关流,关连接
    - 统计 ImportResult
- `src/ExcelImportTool/Services/Import/ProgressInfo.cs`
- `src/ExcelImportTool/Forms/MainForm.ImportTab.Run.cs`(partial)
  - 「开始导入」按钮 → 启动后台 Task;进度条 + 标签(已处理/总数/成功/失败/行秒)
  - 「取消」按钮 → CancellationToken;已提交批保留,未提交回滚,报告已处理数
  - 完成弹汇总对话框:总数/成功/失败/耗时/行秒/失败明细(可导出 CSV,默认存 `%AppData%\ExcelImportTool\reports\`)
  - 导入过程禁止重复点击开始;UI 用 `Invoke` 更新,避免跨线程异常

**验收**
- 小文件(1 万行)端到端导入成功,汇总数字正确
- 中途取消 → 未提交批次回滚,报告已处理数
- 故造错误行(类型不匹配)→ 该批降级逐行,失败行进入明细,其余继续
- 进度条平滑更新,界面不卡死

**注意**
- 进度上报频率:按批上报 + 最多每 100ms 一次 Invoke 合并,避免百万行时 UI 被事件淹没
- 取消时优雅关 `IExcelDataReader`(它可能持有文件句柄)
- 异常分类:连接异常 → 直接中止报告;Excel 读取中断 → 报告已处理数;单行错误 → 收集明细继续

---

## 阶段 6:本地日志与 Tab3

**目标**:导入汇总与失败明细入库,Tab3 可查可归档。

**产出文件**
- `src/ExcelImportTool/Storage/SqliteLogStore.cs`
  - 库路径:`%AppData%\ExcelImportTool\logs\importlog.db`
  - 建表(SQLite):
    - `import_log(id, started_at, finished_at, file_path, table_name, total, success, failed, rows_per_second, summary_text)`
    - `import_failure(id, log_id, row_index, column_name, error, data_preview)`
  - 索引:`import_log.started_at`、`import_failure.log_id`
- `src/ExcelImportTool/Services/LogService.cs`
  - `WriteLog(ImportResult, ImportRequest)` → 事务写入主表 + 失败明细
  - `QueryLogs(DateTime from, DateTime to, int limit) -> List<LogSummaryDto>`
  - `GetFailures(long logId) -> List<FailureDto>`
  - `Archive()`:删除 `started_at < now - retentionDays` 的记录(主表 + 明细级联删);按月命名无意义(SQLite 单库即可),按时间清理即可
  - 失败静默:写日志失败只记 Windows 事件日志/本地 txt,不阻断导入
- `src/ExcelImportTool/Forms/MainForm.LogTab.cs`(partial)
  - 顶部「保留日志」复选框 + 「保留天数」数字框
  - DataGridView:按时间倒序列出导入记录
  - 选中行 → 下方展开失败明细 DataGridView
  - 「立即归档清理」按钮
  - 程序启动时按配置触发一次 Archive

**验收**
- 开关打开 → 每次导入后 SQLite 有对应记录与失败明细
- 关闭开关 → 不写入,但历史仍可查
- 超过保留期的记录在下次启动被清理
- 日志库损坏/锁定 → 不阻断导入,只在 Tab3 显示"日志不可用"

**注意**
- SQLite 并发:导入线程写、UI 线程读,用 WAL 模式(`PRAGMA journal_mode=WAL`)减少锁竞争
- 失败明细可能上万条,查询分页(默认每页 500)
- `data_preview` 截断前 200 字符,避免存全行数据撑爆库

---

## 阶段 7:调优与验收

**目标**:用真实量级数据实测,确认性能指标,固化推荐参数。

**产出**
- `docs/superpowers/reports/2026-08-17-perf-report.md`(实测报告模板)
  - 测试环境:CPU/内存/磁盘/MySQL 版本/网络 RTT
  - 数据矩阵:1万 / 10万 / 100万行 × 字段数(10/30)
  - 指标:总耗时、行/秒、批数、平均批耗时、峰值内存、最大 batch SQL 体积
  - 参数扫描:batchSize ∈ {500,1000,2000,5000,10000}、commitInterval ∈ {2000,5000,10000}
  - 结论:推荐默认值与用户可调范围
- 代码侧:仅做微调(如默认 batch size 按实测调整、SQL 模板缓存、参数对象池化)

**验收**
- 1万行 ≤ 3 秒;10万行 ≤ 30 秒;100万行 ≤ 5 分钟(设计文档目标)
- 若某档不达标,记录原因(索引/网络/服务器)并给出可改进项,不强行 hack 代码

**注意**
- 百万级实测前先用 10 万行做参数扫描,避免每次百万测试耗时
- 监控工具:任务管理器看内存峰值、MySQL `SHOW PROCESSLIST` 看等待事件
- 实测期间不动代码逻辑,只改参数;若需改逻辑另起一次迭代

---

## 通用约定

- 命名:类名 PascalCase,私有字段 `_camelCase`,常量 `UPPER_SNAKE`
- 文件组织:`src/ExcelImportTool/<Area>/<Subdir>/Foo.cs`,partial UI 文件用 `MainForm.<Tab>.<Part>.cs`
- 异常:对外消息用中文,日志用中文,代码内异常 message 用英文 + 关键值
- 编码:所有文件 UTF-8 with BOM(WinForms 源文件兼容性)
- 提交粒度:按阶段提交,commit message 前缀 `[阶段N]`
- 临时脚本/中间产物放 `tmp/`,不进 git
- 不在代码里写死任何真实连接串/密码;测试用串只在本地 `config.local.json`(已 gitignore)

## 风险与回退

| 风险 | 触发 | 缓解 |
|------|------|------|
| MySql.Data 8.x 在 net462 TLS 报错 | 用户环境 SSL 配置 | App.config 显式配置;允许 SslMode=none |
| 百万行 Excel 计数慢 | .xls 格式无快速行数 | 预览不计数,导入时边读边统计;进度条用"已处理数+估算" |
| max_allowed_packet 不足 | batch 太大 | 动态折半重试 |
| SQLite 锁冲突 | 读写同时 | WAL + 短事务 |
| 用户 Excel 首行非列头 | 表头在第 2/3 行 | 预览可见,后续迭代支持"跳过前 N 行"配置 |

## 待用户确认项

以下默认推进,如有异议请在阶段开始前指出:
1. 配置文件位置:`%AppData%\ExcelImportTool\`(跨用户目录,避免工程目录写权限问题)
2. 日志用 SQLite 单库,不再按月分库,仅按时间清理
3. 暂不做"跳过表头行数"配置,默认首行即列头
4. 暂不做 LOAD DATA 模式
5. 单元测试用 xUnit,若环境不便可只保留 BatchInserter/ColumnMapper 的手测脚本
