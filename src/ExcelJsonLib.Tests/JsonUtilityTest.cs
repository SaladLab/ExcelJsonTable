using System;
using System.Linq;
using Xunit;

namespace ExcelJsonLib.Tests
{
    public class JsonUtilityTest
    {
        [Fact]
        public void TestPrettifyJson()
        {
            var j0 = "[\n  1\n ]";
            var j1 = JsonUtility.PrettifyJson(j0);
            Assert.Equal("[ 1 ]", j1);
        }
    }
}
