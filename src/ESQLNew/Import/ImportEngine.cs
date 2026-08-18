using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ESQLNew.Core;
using ESQLNew.Excel;
using MySql.Data.MySqlClient;

namespace ESQLNew.Import
{
    public static class ImportEngine
    {
        public static IList<ColumnInfo> GetTableColumns(string connStr, string table)
        {
            var result = new List<ColumnInfo>();
            using (var conn = new MySqlConnection(connStr))
            {
                conn.Open();
                using (var cmd = new MySqlCommand(
                    "SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, CHARACTER_MAXIMUM_LENGTH FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = @t",
                    conn))
                {
                    cmd.Parameters.AddWithValue("@t", table);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new ColumnInfo
                            {
                                Name = reader.GetString(0),
                                DataType = reader.GetString(1),
                                IsNullable = reader.GetString(2) == "YES",
                                MaxLength = reader.IsDBNull(3) ? 0 : (int)Math.Min(reader.GetInt64(3), int.MaxValue)
                            });
                        }
                    }
                }
            }
            return result;
        }

        public static Task<ImportResult> Run(string connStr, string table, string excelPath,
            int batchSize, int commitEvery, Action<ImportProgress> onProgress, CancellationToken ct)
        {
            IList<string> headers;
            try
            {
                headers = ExcelStreamReader.ReadHeaders(excelPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取 Excel 表头:" + ex.Message, ex);
            }

            IList<ColumnInfo> cols;
            try
            {
                cols = GetTableColumns(connStr, table);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取目标表结构,请检查连接与表名:" + ex.Message, ex);
            }

            if (cols.Count == 0)
                throw new InvalidOperationException("目标表不存在或无列:" + table);

            var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
            var mappings = ColumnMapper.Map(headers, cols, fieldMap);
            var positions = new List<int>();
            for (int i = 0; i < mappings.Count; i++)
                if (mappings[i].Matched)
                    positions.Add(i);

            if (positions.Count == 0)
                throw new InvalidOperationException("Excel 表头与目标表字段无匹配,请检查列名:" + table);

            IEnumerable<object[]> rows = RebuildRows(ExcelStreamReader.ReadRows(excelPath), positions);
            if (AutoIdNeeded(mappings, cols))
            {
                long nextId = GetMaxId(connStr, table) + 1;
                rows = WithAutoId(rows, nextId);
                mappings.Add(new ColumnMapping
                {
                    ExcelColumn = null,
                    TableField = "id",
                    Matched = true,
                    TableColumn = GetColumn(cols, "id")
                });
            }

            var result = new ImportResult();
            BatchInserter.Execute(connStr, table, mappings,
                rows,
                batchSize, commitEvery, onProgress, result, ct);
            return Task.FromResult(result);
        }

        internal static bool AutoIdNeeded(IList<ColumnMapping> mappings, IList<ColumnInfo> cols)
        {
            if (GetColumn(cols, "id") == null) return false;
            foreach (var m in mappings)
                if (m.Matched && string.Equals(m.TableField, "id", StringComparison.OrdinalIgnoreCase))
                    return false;
            return true;
        }

        internal static IEnumerable<object[]> WithAutoId(IEnumerable<object[]> rows, long startId)
        {
            long next = startId;
            foreach (var raw in rows)
            {
                var rebuilt = new object[raw.Length + 1];
                Array.Copy(raw, rebuilt, raw.Length);
                rebuilt[raw.Length] = next++;
                yield return rebuilt;
            }
        }

        private static ColumnInfo GetColumn(IList<ColumnInfo> cols, string name)
        {
            foreach (var c in cols)
                if (string.Equals(c.Name.Trim('`'), name, StringComparison.OrdinalIgnoreCase))
                    return c;
            return null;
        }

        private static long GetMaxId(string connStr, string table)
        {
            using (var conn = new MySqlConnection(connStr))
            {
                conn.Open();
                using (var cmd = new MySqlCommand(
                    "SELECT COALESCE(MAX(id),0) FROM `" + table.Replace("`", "``") + "`", conn))
                {
                    object v = cmd.ExecuteScalar();
                    if (v == null || v == DBNull.Value) return 0;
                    return Convert.ToInt64(v);
                }
            }
        }

        internal static IEnumerable<object[]> RebuildRows(IEnumerable<object[]> rows, IList<int> positions)
        {
            foreach (var raw in rows)
            {
                var rebuilt = new object[positions.Count];
                for (int i = 0; i < positions.Count; i++)
                    rebuilt[i] = raw[positions[i]];
                yield return rebuilt;
            }
        }
    }
}