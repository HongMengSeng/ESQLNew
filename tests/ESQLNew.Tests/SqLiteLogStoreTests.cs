using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using ESQLNew.Import;
using ESQLNew.Logging;
using Xunit;

namespace ESQLNew.Tests
{
    public class SqLiteLogStoreTests : IDisposable
    {
        private readonly string _path;
        private readonly SqLiteLogStore _store;

        public SqLiteLogStoreTests()
        {
            _path = Path.Combine(Path.GetTempPath(), "logstore_" + Guid.NewGuid() + ".db");
            _store = new SqLiteLogStore(_path);
        }

        public void Dispose()
        {
            SQLiteConnection.ClearAllPools();
            try
            {
                File.Delete(_path);
            }
            catch
            {
            }
        }

        [Fact]
        public void LogImport_ThenQuery_ReturnsRowWithFailures()
        {
            var failures = new List<RowFailure>
            {
                new RowFailure { RowNumber = 10, Message = "字段过长" },
                new RowFailure { RowNumber = 5, Message = "格式错误" }
            };
            _store.LogImport(new DateTime(2024, 1, 1, 10, 0, 0), "a.xlsx", "t1", 100, 98, 2, 5000, 100.5, failures);

            var rows = _store.Query(0, 10);
            Assert.Single(rows);
            Assert.Equal("a.xlsx", rows[0][2]);
            Assert.Equal("t1", rows[0][3]);
            Assert.Equal(100L, rows[0][4]);
            Assert.Equal(98L, rows[0][5]);
            Assert.Equal(2L, rows[0][6]);
            Assert.Equal(5000L, rows[0][7]);
            Assert.Equal(100.5, rows[0][8]);

            var got = _store.GetFailures((long)rows[0][0]);
            Assert.Equal(2, got.Count);
            Assert.Equal(5L, got[0][0]);
            Assert.Equal("格式错误", got[0][1]);
            Assert.Equal(10L, got[1][0]);
            Assert.Equal("字段过长", got[1][1]);
        }

        [Fact]
        public void LogImport_WithoutFailures_StoresNone()
        {
            _store.LogImport(DateTime.Now, "b.xlsx", "t2", 10, 10, 0, 100, 10, null);

            var rows = _store.Query(0, 10);
            Assert.Single(rows);
            Assert.Empty(_store.GetFailures((long)rows[0][0]));
        }

        [Fact]
        public void Query_PagesNewestFirst()
        {
            for (int i = 1; i <= 5; i++)
                _store.LogImport(new DateTime(2024, 1, i), "f" + i + ".xlsx", "t", i, i, 0, 1, 1, null);

            var page1 = _store.Query(0, 2);
            Assert.Equal(2, page1.Count);
            Assert.Equal("f5.xlsx", page1[0][2]);
            Assert.Equal("f4.xlsx", page1[1][2]);

            var page3 = _store.Query(2, 2);
            Assert.Single(page3);
            Assert.Equal("f1.xlsx", page3[0][2]);
        }

        [Fact]
        public void PurgeOld_RemovesOnlyExpired()
        {
            _store.LogImport(DateTime.Now.AddDays(-40), "old.xlsx", "t", 1, 1, 0, 1, 1, null);
            _store.LogImport(DateTime.Now, "new.xlsx", "t", 2, 2, 0, 1, 1, null);

            _store.PurgeOld(30);

            var rows = _store.Query(0, 10);
            Assert.Single(rows);
            Assert.Equal("new.xlsx", rows[0][2]);
        }

        [Fact]
        public void PurgeOld_DeletesChildFailures()
        {
            var failures = new List<RowFailure> { new RowFailure { RowNumber = 2, Message = "x" } };
            _store.LogImport(DateTime.Now.AddDays(-40), "old.xlsx", "t", 1, 0, 1, 1, 1, failures);

            _store.PurgeOld(30);

            Assert.Empty(_store.Query(0, 10));
            Assert.Empty(_store.GetFailures(1));
        }
    }
}