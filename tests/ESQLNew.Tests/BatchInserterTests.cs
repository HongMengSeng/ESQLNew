using System.Collections.Generic;
using System.Linq;
using ESQLNew.Import;
using Xunit;

namespace ESQLNew.Tests
{
    public class BatchInserterTests
    {
        [Fact]
        public void BuildInsertSql_PlaceholdersCount()
        {
            var sql = BatchInserter.BuildInsertSql("t", new List<string> { "a", "b" }, 3);
            Assert.Equal("INSERT INTO `t` (`a`,`b`) VALUES (@p0,@p1),(@p2,@p3),(@p4,@p5)", sql);
        }

        [Fact]
        public void BuildInsertSql_BacktickEscape()
        {
            var sql = BatchInserter.BuildInsertSql("we`ird", new List<string> { "a" }, 1);
            Assert.StartsWith("INSERT INTO `we``ird`", sql);
        }

        [Fact]
        public void Split_9RowsBatch4_Yields3()
        {
            var rows = Enumerable.Range(0, 9).Select(i => new object[] { i }).ToList();
            var batches = BatchInserter.Split(rows, 4).ToList();
            Assert.Equal(3, batches.Count);
            Assert.Equal(4, batches[0].Length);
            Assert.Equal(1, batches[2].Length);
        }

        [Fact]
        public void Split_Empty_YieldsNone()
        {
            Assert.Empty(BatchInserter.Split(new List<object[]>(0), 4));
        }
    }
}
