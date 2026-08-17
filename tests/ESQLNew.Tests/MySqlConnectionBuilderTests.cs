using Xunit;

namespace ESQLNew.Tests
{
    public class MySqlConnectionBuilderTests
    {
        [Fact]
        public void Build_IncludesAllParts()
        {
            var cs = ESQLNew.Core.MySqlConnectionBuilder.Build(
                "192.168.201.112", 3306, "root", "pwd", "testdb");
            Assert.Contains("Server=192.168.201.112", cs);
            Assert.Contains("Port=3306", cs);
            Assert.Contains("Uid=root", cs);
            Assert.Contains("Pwd=pwd", cs);
            Assert.Contains("Database=testdb", cs);
            Assert.Contains("SslMode=None", cs);
            Assert.Contains("AllowLoadLocalInfile=false", cs);
            Assert.Contains("CharSet=utf8mb4", cs);
        }
    }
}