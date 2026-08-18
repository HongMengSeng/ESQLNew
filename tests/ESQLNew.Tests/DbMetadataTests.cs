using System;
using System.Collections.Generic;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class DbMetadataTests
    {
        [Theory]
        [InlineData("mysql")]
        [InlineData("information_schema")]
        [InlineData("performance_schema")]
        [InlineData("sys")]
        [InlineData("MYSQL")]
        [InlineData("  sys  ")]
        public void IsSystemDatabase_ReturnsTrue(string name)
        {
            Assert.True(DbMetadata.IsSystemDatabase(name));
        }

        [Theory]
        [InlineData("testdb")]
        [InlineData("esqlnew")]
        [InlineData("sales")]
        [InlineData("")]
        public void IsSystemDatabase_ReturnsFalse(string name)
        {
            Assert.False(DbMetadata.IsSystemDatabase(name));
        }
    }
}