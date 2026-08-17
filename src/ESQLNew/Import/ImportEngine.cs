using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
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

            var mappings = ColumnMapper.Map(headers, cols);
            var positions = new List<int>();
            for (int i = 0; i < mappings.Count; i++)
                if (mappings[i].Matched)
                    positions.Add(i);

            if (positions.Count == 0)
                throw new InvalidOperationException("Excel 表头与目标表字段无匹配,请检查列名:" + table);

            var result = new ImportResult();
            BatchInserter.Execute(connStr, table, mappings,
                RebuildRows(ExcelStreamReader.ReadRows(excelPath), positions),
                batchSize, commitEvery, onProgress, result, ct);
            return Task.FromResult(result);
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