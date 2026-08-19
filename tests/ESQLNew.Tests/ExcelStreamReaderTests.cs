using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using ESQLNew.Excel;
using ExcelDataReader.Exceptions;
using Xunit;

namespace ESQLNew.Tests
{
    public class ExcelStreamReaderTests
    {
        private static bool IsDamage(Exception ex)
        {
            return ex is XmlException || ex is InvalidDataException || ex is ExcelReaderException || ex is IOException;
        }

        [Fact]
        public void SheetNames_ReturnsSheetNames()
        {
            string path = XlsxFixture.Create();
            try
            {
                var names = ExcelStreamReader.SheetNames(path);
                Assert.Contains("Sheet1", names);
            }
            finally
            {
                XlsxFixture.Delete(path);
            }
        }

        [Fact]
        public void ReadHeaders_ReturnsHeaderRow()
        {
            string path = XlsxFixture.Create();
            try
            {
                var headers = ExcelStreamReader.ReadHeaders(path);
                Assert.Equal(3, headers.Count);
                Assert.Equal("担当", headers[0]);
                Assert.Equal("数量", headers[1]);
                Assert.Equal("期限", headers[2]);
            }
            finally
            {
                XlsxFixture.Delete(path);
            }
        }

        [Fact]
        public void ReadRows_SkipsHeaderAndYieldsData()
        {
            string path = XlsxFixture.Create();
            try
            {
                var rows = ExcelStreamReader.ReadRows(path).ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("张三", rows[0][0]);
                Assert.Equal(10.0, (double)rows[0][1]);
                Assert.Equal("2024-01-05", rows[0][2]);
                Assert.Equal("李四", rows[1][0]);
                Assert.Equal(20.0, (double)rows[1][1]);
            }
            finally
            {
                XlsxFixture.Delete(path);
            }
        }

        [Fact]
        public void MissingFile_Throws()
        {
            string missing = Path.Combine(Path.GetTempPath(), "does_not_exist_" + Guid.NewGuid() + ".xlsx");
            Assert.Throws<FileNotFoundException>(() => ExcelStreamReader.ReadHeaders(missing));
        }

        [Fact]
        public void DamagedFile_ThrowsDamageException()
        {
            string fixture = XlsxFixture.Create();
            string damaged = Path.Combine(Path.GetTempPath(), "damaged_" + Guid.NewGuid() + ".xlsx");
            try
            {
                byte[] full = File.ReadAllBytes(fixture);
                byte[] truncated = new byte[(int)(full.Length * 0.6)];
                Array.Copy(full, truncated, truncated.Length);
                File.WriteAllBytes(damaged, truncated);

                var ex = Record.Exception(() => ExcelStreamReader.ReadHeaders(damaged));
                Assert.NotNull(ex);
                Assert.True(IsDamage(ex), "期望损坏类异常,实际: " + ex.GetType().FullName + ": " + ex.Message);
            }
            finally
            {
                XlsxFixture.Delete(fixture);
                XlsxFixture.Delete(damaged);
            }
        }

        public static class XlsxFixture
        {
            public static string Create()
            {
                string path = Path.Combine(Path.GetTempPath(), "fixture_" + Guid.NewGuid() + ".xlsx");
                using (var fs = File.Create(path))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    AddEntry(zip, "[Content_Types].xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                        "</Types>");
                    AddEntry(zip, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "</Relationships>");
                    AddEntry(zip, "xl/workbook.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                        "</workbook>");
                    AddEntry(zip, "xl/_rels/workbook.xml.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                        "</Relationships>");
                    AddEntry(zip, "xl/sharedStrings.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"7\" uniqueCount=\"7\">" +
                        "<si><t>担当</t></si><si><t>数量</t></si><si><t>期限</t></si>" +
                        "<si><t>张三</t></si><si><t>李四</t></si>" +
                        "<si><t>2024-01-05</t></si><si><t>2024-02-05</t></si>" +
                        "</sst>");
                    AddEntry(zip, "xl/worksheets/sheet1.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c></row>" +
                        "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>3</v></c><c r=\"B2\"><v>10</v></c><c r=\"C2\" t=\"s\"><v>5</v></c></row>" +
                        "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>4</v></c><c r=\"B3\"><v>20</v></c><c r=\"C3\" t=\"s\"><v>6</v></c></row>" +
                        "</sheetData></worksheet>");
                }
                return path;
            }

            public static void Delete(string path)
            {
                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }

            public static string CreateMultiSheet()
            {
                string path = Path.Combine(Path.GetTempPath(), "fixture_multi_" + Guid.NewGuid() + ".xlsx");
                using (var fs = File.Create(path))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    AddEntry(zip, "[Content_Types].xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                        "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                        "</Types>");
                    AddEntry(zip, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "</Relationships>");
                    AddEntry(zip, "xl/workbook.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<sheets>" +
                        "<sheet name=\"SheetA\" sheetId=\"1\" r:id=\"rId1\"/>" +
                        "<sheet name=\"SheetB\" sheetId=\"2\" r:id=\"rId2\"/>" +
                        "</sheets>" +
                        "</workbook>");
                    AddEntry(zip, "xl/_rels/workbook.xml.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" +
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                        "</Relationships>");
                    AddEntry(zip, "xl/sharedStrings.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"12\" uniqueCount=\"12\">" +
                        "<si><t>维修单号</t></si><si><t>担当</t></si>" +
                        "<si><t>张三</t></si><si><t>李四</t></si><si><t>王五</t></si>" +
                        "<si><t>赵六</t></si><si><t>孙七</t></si>" +
                        "<si><t>1001</t></si><si><t>1002</t></si><si><t>1003</t></si>" +
                        "<si><t>1004</t></si><si><t>1005</t></si>" +
                        "</sst>");
                    AddEntry(zip, "xl/worksheets/sheet1.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row>" +
                        "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>7</v></c><c r=\"B2\" t=\"s\"><v>2</v></c></row>" +
                        "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>8</v></c><c r=\"B3\" t=\"s\"><v>3</v></c></row>" +
                        "<row r=\"4\"><c r=\"A4\" t=\"s\"><v>9</v></c><c r=\"B4\" t=\"s\"><v>4</v></c></row>" +
                        "</sheetData></worksheet>");
                    AddEntry(zip, "xl/worksheets/sheet2.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>0</v></c></row>" +
                        "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>2</v></c><c r=\"B2\" t=\"s\"><v>7</v></c></row>" +
                        "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>5</v></c><c r=\"B3\" t=\"s\"><v>10</v></c></row>" +
                        "<row r=\"4\"><c r=\"A4\" t=\"s\"><v>6</v></c><c r=\"B4\" t=\"s\"><v>11</v></c></row>" +
                        "</sheetData></worksheet>");
                }
                return path;
            }

            public static void DeleteMultiSheet(string path)
            {
                Delete(path);
            }

            private static void AddEntry(ZipArchive zip, string name, string content)
            {
                var entry = zip.CreateEntry(name);
                using (var writer = new StreamWriter(entry.Open(), Encoding.UTF8))
                    writer.Write(content);
            }
        }
    }
}