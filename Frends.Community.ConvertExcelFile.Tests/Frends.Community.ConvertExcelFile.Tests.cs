using Frends.Community.ConvertExcelFile;
using NUnit.Framework;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;


namespace Frends.Community.ConvertExcelFileTests
{
    [TestFixture]
    public class ExcelConvertTests
    {
        private readonly Input _input = new();
        private readonly Options _options = new();

        // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required).
        // It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/


        [SetUp]
        public void Setup()
        {
            _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../TestData/ExcelTests/In/");
            _options.CsvSeparator = ",";
            _options.ReadOnlyWorkSheetWithName = "";

        }
        [Test]
        public void TestConvertXlsxToCSV()
        {

            // Test converting all worksheets of xlsx file to csv.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToCSV()
        {
            // Test converting all worksheets of xls file to csv.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxToXML()
        {
            // Test converting all worksheets of xlsx file to xml.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToXML()
        {
            // Test converting all worksheets of xls file to xml.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"<workbookworkbook_name=""ExcelTestInput2.xls""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorksheetToXML()
        {
            // Test converting one worksheet of xlsx file to xml.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsOneWorksheetToCSV()
        {
            // Test converting one worksheet of xls file to csv.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = "Foo,Bar,Kanji働,Summa1,2,3,6";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }
        [Test]
        public void TestConvertXlsxToJSON()
        {
            // Test converting all worksheets of xlsx file to JSON.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""1"":[{""A"":""Kissa kuva""},{""B"":""1""},{""C"":""2""},{""D"":""3""}]},{""15"":[{""A"":""Foo""}]},{""16"":[{""B"":""Bar""}]}]}]}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToJSON()
        {
            // Test converting all worksheets of xls file to JSON.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""1"":[{""A"":""Kissa kuva""},{""B"":""1""},{""C"":""2""},{""D"":""3""}]},{""15"":[{""A"":""Foo""}]},{""16"":[{""B"":""Bar""}]}]}]}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }
        [Test]
        public void TestConvertXlsxOneWorksheetToJSON()
        {
            // Test converting one worksheet of xlsx file to JSON.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsOneWorksheetToJSON()
        {
            // Test converting one worksheet of xls file to JSON.
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYY()
        {
            // Test converting worksheet with dates into dd/MM/yyyy format
            _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet2";
            _options.DateFormat = DateFormats.DDMMYYYY;
            _options.ShortDatePattern = false;
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""25/12/2021 0.00.00""},{""B"":""25/02/2021 12.45.41""}, {""C"":""12/05/2020 0.00.00""},{""D"":""30/12/2021 0.00.00""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorkSheetWithDatesMMDDYYYY()
        {
            // Test converting worksheet with dates into MM/dd/yyyy format
            _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            _options.DateFormat = DateFormats.MMDDYYYY;
            _options.ShortDatePattern = false;
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""1""},{""B"":""2""}, {""C"":""3""},{""D"":""4""}]}, {""2"":[{""A"":""12/12/2021 12:00:00AM""},{""B"":""2/25/2021 12:45:41PM""}, {""C"":""5/12/2020 12:00:00AM""},{""D"":""12/12/2021 12:00:00AM""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorkSheetWithDatesYYYYMDD()
        {
            // Test converting worksheet with dates into MM/dd/yyyy format
            _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet2";
            _options.DateFormat = DateFormats.YYYYMMDD;
            _options.ShortDatePattern = false;
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""2021/12/25 0:00:00""},{""B"":""2021/02/25 12:45:41""}, {""C"":""2020/05/12 0:00:00""},{""D"":""2021/12/30 0:00:00""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYYWithShortPattern()
        {
            // Test converting worksheet with dates into dd/MM/yyyy format with ShortTimePattern enabled
            _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
            _options.ReadOnlyWorkSheetWithName = "Sheet2";
            _options.DateFormat = DateFormats.DDMMYYYY;
            _options.ShortDatePattern = true;
            var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
            var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""25/12/2021""},{""B"":""25/02/2021""}, {""C"":""12/05/2020""},{""D"":""30/12/2021""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void ShouldThrowUnknownFileFormatError()
        {
            // Test converting one worksheet of xls file to csv.
            _input.Path = Path.Combine(_input.Path, "UnitTestErrorFile.txt");
            _options.ThrowErrorOnFailure = true;
            Assert.That(() => ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken()),Throws.Exception);
        }

        [Test]
        public void DoNotThrowOnFailure()
        {
            //try to convert a file that does not exist 
            _input.Path = Path.Combine(_input.Path, "thisfiledoesnotexist.txt");
            _options.ThrowErrorOnFailure = false;
            try
            {
                var result = ExcelClass.ConvertExcelFile(_input, _options, new CancellationToken());
                Assert.AreEqual(result.Success, false);
                Assert.AreEqual(result.ResultData, null);
            }
            catch (Exception ex)
            {
                Assert.Fail("This should not happen: " + ex.Message);
            }
        }


    }
}
