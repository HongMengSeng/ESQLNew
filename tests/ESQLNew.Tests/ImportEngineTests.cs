using System.Collections.Generic;
using System.Linq;
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
    }
}