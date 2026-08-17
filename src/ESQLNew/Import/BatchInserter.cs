using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

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

        private static bool TryInsertBatch(MySqlConnection conn, MySqlTransaction tx, string table,
            List<string> fieldNames, List<ColumnMapping> matched, object[] batch,
            ref int batchOk, ref int batchFail, ImportResult result)
        {
            try
            {
                var sql = BuildInsertSql(table, fieldNames, batch.Length);
                using (var cmd = new MySqlCommand(sql, conn, tx))
                {
                    int p = 0;
                    for (int r = 0; r < batch.Length; r++)
                    {
                        var row = (object[])batch[r];
                        for (int i = 0; i < matched.Count; i++)
                        {
                            var col = matched[i].TableColumn ?? new ColumnInfo { Name = matched[i].TableField };
                            cmd.Parameters.AddWithValue("@p" + (p++), ColumnMapper.ConvertValue(row[i], col) ?? DBNull.Value);
                        }
                    }
                    cmd.ExecuteNonQuery();
                }
                batchOk = batch.Length;
                batchFail = 0;
                return true;
            }
            catch (Exception ex)
            {
                result.Failures.Add(new RowFailure { RowNumber = result.Total + 1, Message = ex.Message });
                batchOk = 0;
                batchFail = batch.Length;
                return false;
            }
        }

        private static bool TryInsertRow(MySqlConnection conn, MySqlTransaction tx, string table,
            List<string> fieldNames, List<ColumnMapping> matched, object row, ImportResult result)
        {
            try
            {
                var sql = BuildInsertSql(table, fieldNames, 1);
                using (var cmd = new MySqlCommand(sql, conn, tx))
                {
                    var r = (object[])row;
                    for (int i = 0; i < matched.Count; i++)
                    {
                        var col = matched[i].TableColumn ?? new ColumnInfo { Name = matched[i].TableField };
                        cmd.Parameters.AddWithValue("@p" + i, ColumnMapper.ConvertValue(r[i], col) ?? DBNull.Value);
                    }
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                result.Failures.Add(new RowFailure { RowNumber = result.Total + 1, Message = ex.Message });
                return false;
            }
        }
    }
}