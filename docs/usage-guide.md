# ESQLNew 使用说明与导入逻辑

## 一、使用流程

```
启动 ESQLNew.exe
  │
  ▼
【连接 Tab】
  填 服务器/端口/账号/密码 → 点「测试连接」
     → 成功:自动拉出数据库下拉列表(默认隐藏 mysql 等系统库)
  → 选库(如 impp_i3_qms)→ 自动拉出该库全部表到导入 Tab
  → 点「保存配置」持久化到 %APPDATA%\ESQLNew\config.json
  │
  ▼
【导入 Tab】
  1. 点「选择文件」选 Excel
     → 自动用第一个工作表名预填目标表框
  2. 从表下拉选目标表,或手动输入
  3. 点「匹配预览」→ 显示:
     - 字段映射页:Excel 列 ⇄ 表字段 对应关系 + 匹配状态
     - 样例数据页:前 100 行数据(仅匹配列)
  4. 点「开始导入」→ 进度条 + 已处理/总数/成功/失败 实时计数
     → 完成后弹汇总报告(总数/成功/失败/耗时/行秒/批次数)
     → 有失败可「导出失败明细 CSV」
  │
  ▼
【日志 Tab】
  查看历史导入记录 + 每次导入的失败明细(SQLite 存储)
```

## 二、导入逻辑

```
Excel 文件 (60,391 行 × 81 列)
  │
  ▼ ① 流式读取 (ExcelStreamReader.ReadRows)
  │    逐行读取,不整表加载进内存(百万行也安全)
  │
  ▼ ② 读表头 (ReadHeaders) → 81 个中文列名
  │
  ▼ ③ 读目标表结构 (ImportEngine.GetTableColumns)
  │    查 information_schema.COLUMNS → 89 个英文字段 + 类型
  │
  ▼ ④ 列比对 (ColumnMapper.Map) ★ 关键
  │
  ▼ ⑤ 重建行 (RebuildRows) 按匹配列顺序重排每行
  │
  ▼ ⑥ id 自动生成 (AutoIdNeeded + WithAutoId)
  │    表有 id 主键且 Excel 没提供 → 查 MAX(id)+1 顺序递增
  │
  ▼ ⑦ 批量写入 (BatchInserter.Execute)
  │    多值 INSERT(如一次插 2000 行),每 5000 行提交一次事务
  │    批失败 → 自动降级逐行插入,记录每行失败原因
  │
  ▼ ⑧ 进度回报 → 汇总报告 + 写日志
```

## 三、字段比对逻辑(核心机制)

Excel 表头是中文,数据库字段是英文驼峰,名字完全不同。比对分两层:

### 第 1 层:中文表头 → 英文字段映射表

内置映射字典(见 `src/ESQLNew/Import/ColumnMapper.cs` 的 `FieldMap`):

| Excel 表头 | 表字段 |
|-----------|--------|
| 维修单号 | repair_order_no |
| 担当 | person_in_charge |
| 次数 | order_times |
| 维修次数 | repair_times |
| 机型年度 | model_year |
| 机型年份 | year_label |
| 使用期限(天) | service_life_days |
| 店铺名称 | shop_name |
| 结算金额 | settlement_amount |
| ... | ...(共 81 条,一一对应) |

### 第 2 层:映射命中后找字段类型

映射到字段名后,在表的字段里找到该字段,取出类型与长度限制,用于值转换:

```
匹配到字段 → 拿到 ColumnInfo { Name, DataType, MaxLength }
  │
  ▼ 值转换 (ColumnMapper.ConvertValue)
  │   varchar → 直接字符串(超长则报错)
  │   int/bigint → 解析成整数(Excel 存成数字/文本都能转)
  │   double/decimal → 解析成小数
  │   datetime/date → 三种情况都处理:
  │     ① 已是 DateTime → 直接用
  │     ② Excel 的 OADate 数字(如 43938)→ 转成日期
  │     ③ 文本 "2020-4-17" → 解析
  │
  ▼ 组装 INSERT 参数(参数化查询,防注入)
```

### 匹配失败的兜底

- 映射表查不到的表头 → 退回"按名称直接匹配"(大小写不敏感)
- 仍匹配不上 → 标记"未匹配",预览可见,导入时跳过该列,不影响其它列

## 四、实际效果(60,391 行实测数据)

- 表头 81 列 → 映射表命中 81/81
- 目标表 89 字段 - 8 个系统字段 = 81 业务字段,一一对应
- 60,391 行全部写入成功,零失败,约 1100 行/秒

## 相关文件

- 映射逻辑:`src/ESQLNew/Import/ColumnMapper.cs`
- 导入编排:`src/ESQLNew/Import/ImportEngine.cs`
- 批量写入:`src/ESQLNew/Import/BatchInserter.cs`
- 流式读取:`src/ESQLNew/Excel/ExcelStreamReader.cs`
- 性能报告:`docs/perf-report.md`