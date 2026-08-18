using System;
using System.Collections.Generic;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class ColumnMapStoreTests
    {
        private static string TempPath()
        {
            return Path.Combine(Path.GetTempPath(), "colmap_" + Guid.NewGuid().ToString("N") + ".json");
        }

        [Fact]
        public void BuiltinMap_ContainsRepairTableWith81Entries()
        {
            var map = ColumnMapStore.BuiltinMap;
            Assert.True(map.ContainsKey("sx_qms_repair_info_history"));
            Assert.Equal(81, map["sx_qms_repair_info_history"].Count);
            Assert.Equal("repair_order_no", map["sx_qms_repair_info_history"]["维修单号"]);
        }

        [Fact]
        public void SaveThenLoad_RoundTrip()
        {
            var path = TempPath();
            try
            {
                var maps = new Dictionary<string, Dictionary<string, string>>
                {
                    { "my_table", new Dictionary<string, string> { { "订单号", "order_no" } } }
                };
                ColumnMapStore.Save(path, maps);
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("my_table"));
                Assert.Equal("order_no", loaded["my_table"]["订单号"]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_MissingFile_InitializesWithBuiltin()
        {
            var path = TempPath();
            try
            {
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
                Assert.True(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_CorruptFile_FallsBackToBuiltin()
        {
            var path = TempPath();
            try
            {
                File.WriteAllText(path, "{ not valid json");
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void GetMap_MissingTable_ReturnsEmpty()
        {
            var path = TempPath();
            try
            {
                ColumnMapStore.Save(path, new Dictionary<string, Dictionary<string, string>>());
                var map = ColumnMapStore.GetMap(path, "no_such_table");
                Assert.NotNull(map);
                Assert.Empty(map);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
