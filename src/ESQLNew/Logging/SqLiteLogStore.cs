using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using ESQLNew.Import;

namespace ESQLNew.Logging
{
    public class SqLiteLogStore
    {
        public string DbPath { get; private set; }

        public SqLiteLogStore(string dbPath)
        {
            DbPath = dbPath;
            var dir = Path.GetDirectoryName(dbPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            EnsureSchema();
        }

        public void EnsureSchema()
        {
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS LogImport (Id INTEGER PRIMARY KEY AUTOINCREMENT, Timestamp TEXT, FileName TEXT, TableName TEXT, Total INTEGER, Succeeded INTEGER, Failed INTEGER, ElapsedMs INTEGER, RowsPerSecond REAL)";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS LogFailure (Id INTEGER PRIMARY KEY AUTOINCREMENT, ImportId INTEGER, RowNumber INTEGER, Message TEXT)";
                cmd.ExecuteNonQuery();
            }
        }

        public void LogImport(DateTime ts, string file, string table, long total, long ok, long fail, long ms, double rps, List<RowFailure> failures)
        {
            string stamp = ts.ToString("yyyy-MM-dd HH:mm:ss");
            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                long importId;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "INSERT INTO LogImport (Timestamp, FileName, TableName, Total, Succeeded, Failed, ElapsedMs, RowsPerSecond) VALUES (@ts, @file, @table, @total, @ok, @fail, @ms, @rps)";
                    cmd.Parameters.AddWithValue("@ts", stamp);
                    cmd.Parameters.AddWithValue("@file", file ?? "");
                    cmd.Parameters.AddWithValue("@table", table ?? "");
                    cmd.Parameters.AddWithValue("@total", total);
                    cmd.Parameters.AddWithValue("@ok", ok);
                    cmd.Parameters.AddWithValue("@fail", fail);
                    cmd.Parameters.AddWithValue("@ms", ms);
                    cmd.Parameters.AddWithValue("@rps", rps);
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    importId = (long)cmd.ExecuteScalar();
                }
                if (failures != null && failures.Count > 0)
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = "INSERT INTO LogFailure (ImportId, RowNumber, Message) VALUES (@importId, @row, @msg)";
                        var pId = cmd.Parameters.Add("@importId", DbType.Int64);
                        var pRow = cmd.Parameters.Add("@row", DbType.Int64);
                        var pMsg = cmd.Parameters.Add("@msg", DbType.String);
                        foreach (var f in failures)
                        {
                            pId.Value = importId;
                            pRow.Value = f.RowNumber;
                            pMsg.Value = f.Message ?? "";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                tx.Commit();
            }
        }

        public List<object[]> Query(int page, int pageSize)
        {
            var result = new List<object[]>();
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, Timestamp, FileName, TableName, Total, Succeeded, Failed, ElapsedMs, RowsPerSecond FROM LogImport ORDER BY Id DESC LIMIT @limit OFFSET @offset";
                cmd.Parameters.AddWithValue("@limit", pageSize);
                cmd.Parameters.AddWithValue("@offset", page * pageSize);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new object[]
                        {
                            reader["Id"],
                            reader["Timestamp"],
                            reader["FileName"],
                            reader["TableName"],
                            reader["Total"],
                            reader["Succeeded"],
                            reader["Failed"],
                            reader["ElapsedMs"],
                            reader["RowsPerSecond"]
                        });
                    }
                }
            }
            return result;
        }

        public List<object[]> GetFailures(long importId)
        {
            var result = new List<object[]>();
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT RowNumber, Message FROM LogFailure WHERE ImportId = @importId ORDER BY RowNumber";
                cmd.Parameters.AddWithValue("@importId", importId);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new object[] { reader["RowNumber"], reader["Message"] });
                    }
                }
            }
            return result;
        }

        public void PurgeOld(int days)
        {
            string cutoff = DateTime.Now.AddDays(-days).ToString("yyyy-MM-dd HH:mm:ss");
            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "DELETE FROM LogFailure WHERE ImportId IN (SELECT Id FROM LogImport WHERE Timestamp < @cutoff)";
                    cmd.Parameters.AddWithValue("@cutoff", cutoff);
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "DELETE FROM LogImport WHERE Timestamp < @cutoff";
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();
            }
        }

        private SQLiteConnection Open()
        {
            var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;");
            conn.Open();
            return conn;
        }
    }
}