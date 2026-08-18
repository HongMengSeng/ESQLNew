using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;

namespace ESQLNew.Core
{
    public static class DbMetadata
    {
        private static readonly string[] SystemDatabases = { "mysql", "information_schema", "performance_schema", "sys" };

        public static bool IsSystemDatabase(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            var trimmed = name.Trim();
            foreach (var s in SystemDatabases)
                if (string.Equals(trimmed, s, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        public static IList<string> GetDatabases(string connStr)
        {
            var result = new List<string>();
            try
            {
                using (var conn = new MySqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand("SHOW DATABASES", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            result.Add(reader.GetString(0));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取数据库列表:" + ex.Message, ex);
            }
            return result;
        }

        public static IList<string> GetTables(string connStr, string database)
        {
            var result = new List<string>();
            try
            {
                using (var conn = new MySqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand(
                        "SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = @db AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME",
                        conn))
                    {
                        cmd.Parameters.AddWithValue("@db", database);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                result.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取表列表:" + ex.Message, ex);
            }
            return result;
        }
    }
}