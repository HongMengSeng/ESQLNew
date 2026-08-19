using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ESQLNew.Core;
using ESQLNew.Import;
using Xunit;

namespace ESQLNew.Tests
{
    public class ImportEngineTests
    {
        [Fact]
        public void RebuildRows_OrdersByMatchedPositions()
        {
            var positions = new List<int> { 2, 0, 3 };
            var rows = new List<object[]> { new object[] { "a", "b", "c", "d" } };
            var rebuilt = ImportEngine.RebuildRows(rows, positions).ToList();
            Assert.Equal(3, rebuilt[0].Length);
            Assert.Equal("c", rebuilt[0][0]);
            Assert.Equal("a", rebuilt[0][1]);
            Assert.Equal("d", rebuilt[0][2]);
        }

        [Fact]
        public void WithAutoId_AppendsSequentialIds()
        {
            var rows = new List<object[]> { new object[] { "a", 1 }, new object[] { "b", 2 } };
            var result = ImportEngine.WithAutoId(rows, 100).ToList();
            Assert.Equal(3, result[0].Length);
            Assert.Equal(100L, result[0][2]);
            Assert.Equal("a", result[0][0]);
            Assert.Equal(101L, result[1][2]);
            Assert.Equal("b", result[1][0]);
        }

        [Fact]
        public void AutoIdNeeded_TrueWhenIdColumnUnmapped()
        {
            var headers = new List<string> { "维修单号" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "id", DataType = "bigint", IsNullable = false, MaxLength = 0 }
            };
            var mappings = ColumnMapper.Map(headers, cols);
            Assert.True(ImportEngine.AutoIdNeeded(mappings, cols));
        }

        [Fact]
        public void Run_UsesTableSpecificExternalMap()
        {
            var path = Path.Combine(Path.GetTempPath(), "impmap_" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                ColumnMapStore.Save(path, new Dictionary<string, Dictionary<string, string>>
                {
                    { "custom_table", new Dictionary<string, string> { { "自定义表头", "custom_field" } } }
                });
                var map = ColumnMapStore.GetMap(path, "custom_table");
                Assert.Equal("custom_field", map["自定义表头"]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void AutoIdNeeded_FalseWhenIdAlreadyMapped()
        {
            var headers = new List<string> { "id" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "id", DataType = "bigint", IsNullable = false, MaxLength = 0 }
            };
            var mappings = ColumnMapper.Map(headers, cols);
            Assert.False(ImportEngine.AutoIdNeeded(mappings, cols));
        }

        [Fact]
        public void MergeSheets_AlignsAndDedups()
        {
            string path = ESQLNew.Tests.ExcelStreamReaderTests.XlsxFixture.CreateMultiSheet();
            try
            {
                var positions = new List<int> { 0, 1 };
                var canonical = new List<ColumnMapping>
                {
                    new ColumnMapping { ExcelColumn = "维修单号", TableField = "repair_order_no", Matched = true },
                    new ColumnMapping { ExcelColumn = "担当", TableField = "person_in_charge", Matched = true }
                };
                var cols = new List<ColumnInfo>
                {
                    new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                    new ColumnInfo { Name = "person_in_charge", DataType = "varchar", IsNullable = true, MaxLength = 50 }
                };
                var fieldMap = new Dictionary<string, string>
                {
                    { "维修单号", "repair_order_no" },
                    { "担当", "person_in_charge" }
                };
                var merged = ImportEngine.MergeSheets(path, new List<string> { "SheetA", "SheetB" },
                    positions, canonical, cols, fieldMap, "维修单号").ToList();
                Assert.Equal(5, merged.Count);
                Assert.Equal("1001", merged[0][0] == null ? "" : merged[0][0].ToString());
                Assert.Equal("张三", merged[0][1] == null ? "" : merged[0][1].ToString());
            }
            finally
            {
                ESQLNew.Tests.ExcelStreamReaderTests.XlsxFixture.DeleteMultiSheet(path);
            }
        }
    }
}