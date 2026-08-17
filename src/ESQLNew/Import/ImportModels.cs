using System;

namespace ESQLNew.Import
{
    public class ColumnInfo
    {
        public string Name { get; set; }
        public string DataType { get; set; }
        public bool IsNullable { get; set; }
        public int MaxLength { get; set; }
    }

    public class ColumnMapping
    {
        public string ExcelColumn { get; set; }
        public string TableField { get; set; }
        public bool Matched { get; set; }
    }

    public class RowFailure
    {
        public long RowNumber { get; set; }
        public string Message { get; set; }
    }

    public class ImportResult
    {
        public long Total { get; set; }
        public long Succeeded { get; set; }
        public long Failed { get; set; }
        public System.Collections.Generic.List<RowFailure> Failures { get; set; } =
            new System.Collections.Generic.List<RowFailure>();
        public TimeSpan Elapsed { get; set; }
        public double RowsPerSecond { get; set; }
        public int BatchCount { get; set; }
    }

    public class ImportProgress
    {
        public long Processed { get; set; }
        public long Total { get; set; }
        public long Succeeded { get; set; }
        public long Failed { get; set; }
    }
}