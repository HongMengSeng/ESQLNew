namespace ESQLNew.Core
{
    public static class MySqlConnectionBuilder
    {
        public static string Build(string server, int port, string user, string password, string database)
        {
            return string.Format(
                "Server={0};Port={1};Uid={2};Pwd={3};Database={4};SslMode=None;AllowLoadLocalInfile=false;CharSet=utf8mb4;",
                server, port, user, password, database);
        }
    }
}