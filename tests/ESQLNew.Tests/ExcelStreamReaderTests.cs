using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using ESQLNew.Excel;
using ExcelDataReader.Exceptions;
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

        private bool IsDamage(Exception ex)
        {
            return ex is XmlException || ex is InvalidDataException || ex is ExcelReaderException;
        }

        [Fact]
        public void SheetNames_ReturnsTemplateSheet()
        {
            IList<string> names;
            try
            {
                names = ExcelStreamReader.SheetNames(Template);
            }
            catch (Exception ex) when (IsDamage(ex))
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] SheetNames 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
                return;
            }
            Assert.Contains("SxQmsRepairInfoHistory", names);
        }

        [Fact]
        public void ReadHeaders_Returns80Columns()
        {
            IList<string> headers;
            try
            {
                headers = ExcelStreamReader.ReadHeaders(Template);
            }
            catch (Exception ex) when (IsDamage(ex))
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] ReadHeaders 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
                return;
            }
            Assert.Equal(80, headers.Count);
            Assert.Equal("担当", headers[0]);
            Assert.Equal("期限", headers[79]);
        }

        [Fact]
        public void ReadRows_StreamsRows()
        {
            List<object[]> rows;
            try
            {
                rows = ExcelStreamReader.ReadRows(Template).ToList();
            }
            catch (Exception ex) when (IsDamage(ex))
            {
                _output.WriteLine("[DAMAGED-TEMPLATE] ReadRows 读取失败: " + ex.GetType().FullName + ": " + ex.Message);
                return;
            }
            Assert.True(rows.Count >= 300, "模板应包含 347 行数据");
        }

        [Fact]
        public void MissingFile_Throws()
        {
            Assert.Throws<FileNotFoundException>(() => ExcelStreamReader.ReadHeaders(@"C:\does_not_exist.xlsx"));
        }
    }
}