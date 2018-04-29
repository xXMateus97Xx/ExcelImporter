using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExcelImporter.Tests
{
    [TestFixture]
    public class ExcelImporterTest
    {
        [SetUp]
        public void Init()
        {
            ImporterMapper<Person>.Create()
                .AddField(x => x.Age, "Age", 3, "20")
                .AddField(x => x.Birthday, "Birthday Date", 4, "07/25/1997")
                .AddField(x => x.HasJob, "Has Job", 5, "True")
                .AddField(x => x.HasPet, "Has Pet", 7, "False")
                .AddField(x => x.Name, "Name", 1, "John Smith")
                .AddField(x => x.Nickname, "Nickname", 2, "Jojo")
                .AddField(x => x.PetAge, "Pet Age", 9, "5")
                .AddField(x => x.PetName, "Pet Name", 8, "Rex")
                .AddField(x => x.Weight, "Weight", 6, "85")
                .AddExtraField("Marriage", typeof(bool), "Is Marriage", 11, "True")
                .AddExtraField("ChildrenQuantity", typeof(int), "Children(s)", 10, "2")
                .SetWorksheetName("People")
                .Map();
        }

        [Test]
        public void ShouldMapPerson()
        {
            ImporterMapper<Person>.Create()
                .AddField(x => x.Age, "Age", 3, "20")
                .AddField(x => x.Birthday, "Birthday Date", 4, "07/25/1997")
                .AddField(x => x.HasJob, "Has Job", 5, "True")
                .AddField(x => x.HasPet, "Has Pet", 7, "False")
                .AddField(x => x.Name, "Name", 1, "John Smith")
                .AddField(x => x.Nickname, "Nickname", 2, "Jojo")
                .AddField(x => x.PetAge, "Pet Age", 9, "5")
                .AddField(x => x.PetName, "Pet Name", 8, "Rex")
                .AddField(x => x.Weight, "Weight", 6, "85")
                .AddExtraField("Marriage", typeof(bool), "Is Marriage", 11, "True")
                .AddExtraField("ChildrenQuantity", typeof(int), "Children(s)", 10, "2")
                .SetWorksheetName("People")
                .Map();

            Assert.True(true);
        }

        [Test]
        public void ShouldGetSampleExcel()
        {
            var importer = new Importer<Person>();

            var file = importer.GetSampleExcelFile();

            File.WriteAllBytes("C:\\ExcelImporterTests\\PersonSample.xlsx", file);

            Assert.True(true);
        }

        [Test]
        public void ShouldImportWithoutErrors()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            var file = File.ReadAllBytes("C:\\ExcelImporterTests\\OkTest.xlsx");
            var importer = new Importer<Person>();

            var items = importer.Import(file);

            Assert.False(items.Any(x => x.HasError));
        }

        [Test]
        public void ShouldImportUsingDifferentCulture()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("pt-BR");

            var file = File.ReadAllBytes("C:\\ExcelImporterTests\\CultureTest.xlsx");
            var importer = new Importer<Person>();

            var items = importer.Import(file);

            Assert.False(items.Any(x => x.HasError));
        }
    }
}
