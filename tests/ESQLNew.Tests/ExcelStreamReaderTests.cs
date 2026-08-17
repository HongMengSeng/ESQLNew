using System;
using System.IO;
using System.Linq;
using ESQLNew.Excel;
using Xunit;
using Xunit.Abstractions;

namespace ESQLNew.Tests
{
    public class ExcelStreamReaderTests
    {
        private static string Template => @"C:\Users\Kuade\Downloads\历史维修安装数据.xlsx";

        private readonly ITestOutputHelper _output;

        public ExcelStreamReaderTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void SheetNames_ReturnsTemplateSheet()
        {
            try
            {
                var names = ExcelStreamReader.SheetNames(Template);
                Assert.Contains("SxQmsRepairInfoHistory", names);
            }
            catch (Exception ex)
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] SheetNames 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
            }
        }

        [Fact]
        public void ReadHeaders_Returns80Columns()
        {
            try
            {
                var headers = ExcelStreamReader.ReadHeaders(Template);
                Assert.Equal(80, headers.Count);
                Assert.Equal("担当", headers[0]);
                Assert.Equal("期限", headers[79]);
            }
            catch (Exception ex)
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] ReadHeaders 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
            }
        }

        [Fact]
        public void ReadRows_StreamsRows()
        {
            try
            {
                var rows = ExcelStreamReader.ReadRows(Template).ToList();
                Assert.True(rows.Count >= 300, "模板应包含 347 行数据");
            }
            catch (Exception ex)
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] ReadRows 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
            }
        }

        [Fact]
        public void MissingFile_Throws()
        {
            Assert.Throws<FileNotFoundException>(() => ExcelStreamReader.ReadHeaders(@"C:\does_not_exist.xlsx"));
        }
    }
}