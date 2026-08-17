using System.Collections.Generic;
using ESQLNew.Import;
using Xunit;

namespace ESQLNew.Tests
{
    public class ColumnMapperTests
    {
        [Fact]
        public void Map_MatchesByName_CaseInsensitive()
        {
            var headers = new List<string> { "担当", "次数", "不存在的列" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "担当", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "次数", DataType = "int", IsNullable = false, MaxLength = 0 }
            };
            var map = ColumnMapper.Map(headers, cols);
            Assert.Equal(3, map.Count);
            Assert.True(map[0].Matched);
            Assert.Equal("担当", map[0].TableField);
            Assert.False(map[2].Matched);
        }

        [Fact]
        public void ConvertValue_EmptyToNull()
        {
            var col = new ColumnInfo { Name = "a", DataType = "varchar", IsNullable = true, MaxLength = 50 };
            Assert.Null(ColumnMapper.ConvertValue("", col));
        }

        [Fact]
        public void ConvertValue_DateOADateToDateTime()
        {
            var col = new ColumnInfo { Name = "d", DataType = "datetime", IsNullable = true, MaxLength = 0 };
            var dt = ColumnMapper.ConvertValue(new System.DateTime(2020, 4, 17), col);
            Assert.Equal(new System.DateTime(2020, 4, 17), dt);
        }

        [Fact]
        public void ConvertValue_NumericStringToLong()
        {
            var col = new ColumnInfo { Name = "n", DataType = "int", IsNullable = true, MaxLength = 0 };
            Assert.Equal(123L, ColumnMapper.ConvertValue("123", col));
        }
    }
}