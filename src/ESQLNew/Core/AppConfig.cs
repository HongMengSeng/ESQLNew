using System;
using System.IO;
using Newtonsoft.Json;

namespace ESQLNew.Core
{
    public class AppConfig
    {
        public string Server { get; set; }
        public int Port { get; set; } = 3306;
        public string User { get; set; }
        public string Password { get; set; }
        public string Database { get; set; }
        public string RawConnectionString { get; set; }
        public bool UseRaw { get; set; }
        public int BatchSize { get; set; } = 2000;
        public int CommitEvery { get; set; } = 5000;
        public bool KeepLogs { get; set; } = true;
        public int LogRetentionDays { get; set; } = 30;
        public bool ShowSystemDatabases { get; set; } = false;

        public static AppConfig Load(string path)
        {
            try
            {
                if (File.Exists(path))
                    return JsonConvert.DeserializeObject<AppConfig>(File.ReadAllText(path));
            }
            catch
            {
            }
            return new AppConfig();
        }

        public void Save(string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}