using AutoFixture;
using SpreadsheetFluent.Test.Models;
using SpreadsheetFluent.Test.WorkBooks;
using System;
using Xunit;
using System.Linq;
using System.IO;

namespace SpreadsheetFluent.Test
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            var fixture = new Fixture();
            var persons = fixture.CreateMany<Person>(100).ToList();

            var workbook = new EmployeeWorkbook();

            var result = workbook.Generate(persons);

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Test2.xlsx");

            if (File.Exists(path)) File.Delete(path);

            File.WriteAllBytes(path, result);

            Assert.True(result.Length > 0);
        }
    }
}
