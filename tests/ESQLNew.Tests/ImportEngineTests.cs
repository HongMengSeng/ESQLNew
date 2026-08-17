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
    }
}