using System;
using System.IO;

namespace ESQLNew.Core
{
    public static class AppPaths
    {
        private static string _configDir;
        private static string _configFile;
        private static string _columnMapsFile;
        private static string _logDbFile;
        private static string _legacyDir;

        public static string ConfigDir { get { return _configDir; } }
        public static string ConfigFile { get { return _configFile; } }
        public static string ColumnMapsFile { get { return _columnMapsFile; } }
        public static string LogDbFile { get { return _logDbFile; } }
        public static string LegacyDir { get { return _legacyDir; } }

        public static void Initialize(string baseDirectory, string appPathTxtPath, string legacyDir)
        {
            string dir = baseDirectory;
            try
            {
                if (File.Exists(appPathTxtPath))
                {
                    string line = null;
                    using (var r = new StreamReader(appPathTxtPath))
                        line = r.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                        dir = Path.GetFullPath(Path.Combine(baseDirectory, line.Trim()));
                }
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
            catch
            {
                dir = baseDirectory;
            }
            _configDir = dir;
            _configFile = Path.Combine(dir, "config.json");
            _columnMapsFile = Path.Combine(dir, "column-maps.json");
            _logDbFile = Path.Combine(dir, "logs.db");
            _legacyDir = legacyDir;
        }

        public static void MigrateFromLegacy()
        {
            try
            {
                if (string.IsNullOrEmpty(_legacyDir) || !Directory.Exists(_legacyDir)) return;
                string[] files = { "config.json", "column-maps.json", "logs.db" };
                foreach (var f in files)
                {
                    try
                    {
                        string src = Path.Combine(_legacyDir, f);
                        string dst = Path.Combine(_configDir, f);
                        if (File.Exists(src) && !File.Exists(dst))
                            File.Copy(src, dst);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        static AppPaths()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            Initialize(baseDir, Path.Combine(baseDir, "app-path.txt"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ESQLNew"));
        }
    }
}