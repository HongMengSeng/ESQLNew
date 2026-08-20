using System;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class AppPathsTests : IDisposable
    {
        private readonly string _tmp;
        public AppPathsTests()
        {
            _tmp = Path.Combine(Path.GetTempPath(), "apppaths_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tmp);
        }
        public void Dispose()
        {
            try { Directory.Delete(_tmp, true); } catch { }
        }

        [Fact]
        public void ConfigDir_DefaultsToBaseDirectory()
        {
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(_tmp, AppPaths.ConfigDir);
        }

        [Fact]
        public void ConfigDir_UsesAppPathTxtWhenPresent()
        {
            var custom = Path.Combine(_tmp, "custom");
            File.WriteAllText(Path.Combine(_tmp, "app-path.txt"), custom);
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.GetFullPath(custom), Path.GetFullPath(AppPaths.ConfigDir));
            Assert.True(Directory.Exists(AppPaths.ConfigDir));
        }

        [Fact]
        public void ConfigDir_AppPathTxtEmpty_FallsBackToBase()
        {
            File.WriteAllText(Path.Combine(_tmp, "app-path.txt"), "   ");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.GetFullPath(_tmp), Path.GetFullPath(AppPaths.ConfigDir));
        }

        [Fact]
        public void FileProperties_PointIntoConfigDir()
        {
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), Path.Combine(_tmp, "legacy"));
            Assert.Equal(Path.Combine(_tmp, "config.json"), AppPaths.ConfigFile);
            Assert.Equal(Path.Combine(_tmp, "column-maps.json"), AppPaths.ColumnMapsFile);
            Assert.Equal(Path.Combine(_tmp, "logs.db"), AppPaths.LogDbFile);
        }

        [Fact]
        public void MigrateFromLegacy_CopiesMissingFiles()
        {
            var legacy = Path.Combine(_tmp, "legacy");
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "config.json"), "cfg");
            File.WriteAllText(Path.Combine(legacy, "logs.db"), "log");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), legacy);
            AppPaths.MigrateFromLegacy();
            Assert.True(File.Exists(Path.Combine(_tmp, "config.json")));
            Assert.True(File.Exists(Path.Combine(_tmp, "logs.db")));
            Assert.Equal("cfg", File.ReadAllText(Path.Combine(_tmp, "config.json")));
        }

        [Fact]
        public void MigrateFromLegacy_DoesNotOverwriteExisting()
        {
            var legacy = Path.Combine(_tmp, "legacy");
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "config.json"), "legacy");
            File.WriteAllText(Path.Combine(_tmp, "config.json"), "new");
            AppPaths.Initialize(_tmp, Path.Combine(_tmp, "app-path.txt"), legacy);
            AppPaths.MigrateFromLegacy();
            Assert.Equal("new", File.ReadAllText(Path.Combine(_tmp, "config.json")));
        }
    }
}
