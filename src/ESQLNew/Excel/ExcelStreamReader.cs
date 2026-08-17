using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ESQLNew.Excel
{
    public static class ExcelStreamReader
    {
        public static IList<string> SheetNames(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var names = new List<string>();
                do
                {
                    names.Add(reader.Name);
                } while (reader.NextResult());
                return names;
            }
        }

        public static IList<string> ReadHeaders(string path, string sheetName = null)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    var name = reader.Name;
                    var header = new List<string>();
                    if (sheetName == null || string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                                header.Add(reader.GetValue(i) == null ? "" : reader.GetValue(i).ToString());
                        }
                        return header;
                    }
                } while (reader.NextResult());
            }
            return new List<string>();
        }

        public static IEnumerable<object[]> ReadRows(string path, string sheetName = null)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    var name = reader.Name;
                    if (sheetName == null || string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        bool first = true;
                        while (reader.Read())
                        {
                            if (first) { first = false; continue; }
                            var values = new object[reader.FieldCount];
                            for (int i = 0; i < reader.FieldCount; i++)
                                values[i] = reader.GetValue(i);
                            yield return values;
                        }
                        yield break;
                    }
                } while (reader.NextResult());
            }
        }
    }
}