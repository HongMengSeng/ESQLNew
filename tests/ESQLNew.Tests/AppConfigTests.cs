using System;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class AppConfigTests
    {
        [Fact]
        public void SaveLoad_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = new AppConfig
            {
                Server = "192.168.201.112",
                Port = 3306,
                User = "root",
                Password = "pwd",
                Database = "testdb",
                RawConnectionString = "",
                UseRaw = false,
                BatchSize = 2000,
                CommitEvery = 5000,
                KeepLogs = true,
                LogRetentionDays = 30
            };
            cfg.Save(path);
            var loaded = AppConfig.Load(path);
            Assert.Equal("192.168.201.112", loaded.Server);
            Assert.Equal(3306, loaded.Port);
            Assert.Equal("root", loaded.User);
            Assert.Equal("pwd", loaded.Password);
            Assert.Equal("testdb", loaded.Database);
            Assert.Equal(2000, loaded.BatchSize);
            Assert.Equal(true, loaded.KeepLogs);
            File.Delete(path);
        }

        [Fact]
        public void Load_Missing_ReturnsDefaults()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_missing_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = AppConfig.Load(path);
            Assert.Equal(3306, cfg.Port);
            Assert.Equal(2000, cfg.BatchSize);
            Assert.Equal(5000, cfg.CommitEvery);
            Assert.True(cfg.KeepLogs);
        }

        [Fact]
        public void Default_ShowSystemDatabases_False()
        {
            var path = Path.Combine(Path.GetTempPath(), "esqlnew_sysdb_" + Guid.NewGuid().ToString("N") + ".json");
            var cfg = AppConfig.Load(path);
            Assert.False(cfg.ShowSystemDatabases);
        }
    }
}