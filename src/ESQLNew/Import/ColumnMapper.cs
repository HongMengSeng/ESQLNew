using System;
using System.Collections.Generic;
using System.Globalization;
using ESQLNew.Core;

namespace ESQLNew.Import
{
    public static class ColumnMapper
    {
        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)
        {
            Dictionary<string, string> builtin;
            ColumnMapStore.BuiltinMap.TryGetValue("sx_qms_repair_info_history", out builtin);
            return Map(excelHeaders, tableColumns, builtin ?? new Dictionary<string, string>());
        }

        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns,
            IDictionary<string, string> fieldMap)
        {
            var result = new List<ColumnMapping>(excelHeaders.Count);
            var usedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in excelHeaders)
            {
                ColumnInfo matched = null;
                string mappedField = null;
                string key = h.Trim();
                if (fieldMap != null && fieldMap.TryGetValue(key, out mappedField))
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), mappedField, StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                if (matched == null)
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), h.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                if (matched != null && !usedFields.Add(matched.Name.Trim('`')))
                    matched = null;
                result.Add(new ColumnMapping
                {
                    ExcelColumn = h,
                    TableField = matched != null ? matched.Name : null,
                    Matched = matched != null,
                    TableColumn = matched
                });
            }
            return result;
        }

        public static object ConvertValue(object raw, ColumnInfo col)
        {
            if (raw == null) return null;
            string s = raw as string;
            if (s != null && s.Trim().Length == 0) return null;

            string type = (col.DataType ?? "").ToLowerInvariant();

            if (type == "datetime" || type == "timestamp" || type == "date" || type == "time")
            {
                var dt = raw as DateTime?;
                if (dt.HasValue) return dt.Value;
                var oa = raw as double?;
                if (!oa.HasValue) oa = raw as float?;
                if (oa.HasValue)
                {
                    try
                    {
                        return DateTime.FromOADate(oa.Value);
                    }
                    catch (ArgumentException)
                    {
                        throw new FormatException("无法解析日期: " + raw);
                    }
                }
                if (s != null)
                {
                    DateTime parsed;
                    if (DateTime.TryParse(s.Trim(), out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析日期: " + s);
            }

            if (type == "int" || type == "bigint" || type == "smallint" || type == "tinyint" || type == "mediumint")
            {
                var n = raw as int?;
                if (n.HasValue) return (long)n.Value;
                var l = raw as long?;
                if (l.HasValue) return l.Value;
                var d = raw as double?;
                if (d.HasValue) return (long)d.Value;
                if (s != null)
                {
                    long parsed;
                    if (long.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                    double dd;
                    if (double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out dd))
                        return (long)dd;
                }
                throw new FormatException("无法解析整数: " + s);
            }

            if (type == "decimal" || type == "double" || type == "float")
            {
                var d = raw as double?;
                if (d.HasValue) return Convert.ToDecimal(d.Value);
                var dec = raw as decimal?;
                if (dec.HasValue) return dec.Value;
                if (s != null)
                {
                    decimal parsed;
                    if (decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析数字: " + s);
            }

            if (col.MaxLength > 0 && s != null && s.Length > col.MaxLength)
                throw new FormatException("超出字段长度 " + col.MaxLength + ": " + s);

            return s ?? raw.ToString();
        }
    }
}