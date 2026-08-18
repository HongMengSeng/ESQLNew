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
        public void Map_ChineseHeaderMapsToEnglishField()
        {
            var headers = new List<string> { "维修单号", "担当", "机型年度", "无映射列" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "person_in_charge", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "model_year", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var map = ColumnMapper.Map(headers, cols);
            Assert.Equal(4, map.Count);
            Assert.True(map[0].Matched);
            Assert.Equal("repair_order_no", map[0].TableField);
            Assert.True(map[1].Matched);
            Assert.Equal("person_in_charge", map[1].TableField);
            Assert.True(map[2].Matched);
            Assert.Equal("model_year", map[2].TableField);
            Assert.False(map[3].Matched);
        }

        [Fact]
        public void Map_MappingWinsOverNameCollision()
        {
            var headers = new List<string> { "担当" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "person_in_charge", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var map = ColumnMapper.Map(headers, cols);
            Assert.True(map[0].Matched);
            Assert.Equal("person_in_charge", map[0].TableField);
        }

        [Fact]
        public void Map_ExternalMapTakesPriority()
        {
            var headers = new List<string> { "维修单号" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 },
                new ColumnInfo { Name = "repair_order_no", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var fieldMap = new Dictionary<string, string> { { "维修单号", "repair_order_no" } };
            var map = ColumnMapper.Map(headers, cols, fieldMap);
            Assert.True(map[0].Matched);
            Assert.Equal("repair_order_no", map[0].TableField);
        }

        [Fact]
        public void Map_ExternalMapMiss_FallsBackToName()
        {
            var headers = new List<string> { "自定义列" };
            var cols = new List<ColumnInfo>
            {
                new ColumnInfo { Name = "自定义列", DataType = "varchar", IsNullable = true, MaxLength = 50 }
            };
            var fieldMap = new Dictionary<string, string> { { "维修单号", "repair_order_no" } };
            var map = ColumnMapper.Map(headers, cols, fieldMap);
            Assert.True(map[0].Matched);
            Assert.Equal("自定义列", map[0].TableField);
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

        [Fact]
        public void ConvertValue_OADateDoubleToDateTime()
        {
            var col = new ColumnInfo { Name = "d", DataType = "datetime", IsNullable = true, MaxLength = 0 };
            var dt = ColumnMapper.ConvertValue(43938.0, col);
            Assert.Equal(new System.DateTime(2020, 4, 17), dt);
        }
    }
}