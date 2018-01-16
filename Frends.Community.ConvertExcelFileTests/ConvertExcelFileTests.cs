using Frends.Community.ConvertExcelFile;
using NUnit.Framework;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Frends.Community.ConvertExcelFileTests
{
    [TestFixture]
    public class ExcelConvertTests
    {
        private readonly Input _input = new Input();
        private readonly Options _options = new Options();

        // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required)
        // It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/


        [SetUp]
        public void Setup()
        {
            _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\TestData\ExcelTests\In\");
            _options.CsvSeparator = ",";
            _options.ReadOnlyWorkSheetWithName = "";

        }
        [Test]
        public void TestConvertXlsxToCSV()
        {

            // Test converting all worksheets of xlsx file to csv
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToCSV()
        {
            // Test converting all worksheets of xls file to csv 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxToXML()
        {
            // Test converting all worksheets of xlsx file to xml 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToXML()
        {
            // Test converting all worksheets of xls file to xml 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"<workbookworkbook_name=""ExcelTestInput2.xls""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsxOneWorksheetToXML()
        {
            // Test converting one worksheet of xlsx file to xml 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet></workbook>";
            Assert.That(Regex.Replace(result.ToXml(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsOneWorksheetToCSV()
        {
            // Test converting one worksheet of xls file to csv 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = "Foo,Bar,Kanji働,Summa1,2,3,6";
            Assert.That(Regex.Replace(result.ToCsv(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }
        [Test]
        public void TestConvertXlsxToJSON()
        {
            // Test converting all worksheets of xlsx file to JSON
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"{""workbook"":{""@workbook_name"":""ExcelTestInput1.xlsx"",""worksheet"":[{""@worksheet_name"":""Sheet1"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Foo""},{""@column_header"":""B"",""#text"":""Bar""},{""@column_header"":""C"",""#text"":""Kanji働""},{""@column_header"":""D"",""#text"":""Summa""}]},{""@row_header"":""2"",""column"":[{""@column_header"":""A"",""#text"":""1""},{""@column_header"":""B"",""#text"":""2""},{""@column_header"":""C"",""#text"":""3""},{""@column_header"":""D"",""#text"":""6""}]}]},{""@worksheet_name"":""OmituinenNimi"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Kissakuva""},{""@column_header"":""B"",""#text"":""1""},{""@column_header"":""C"",""#text"":""2""},{""@column_header"":""D"",""#text"":""3""}]},{""@row_header"":""15"",""column"":{""@column_header"":""A"",""#text"":""Foo""}},{""@row_header"":""16"",""column"":{""@column_header"":""B"",""#text"":""Bar""}}]}]}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsToJSON()
        {
            // Test converting all worksheets of xls file to JSON
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"{""workbook"":{""@workbook_name"":""ExcelTestInput2.xls"",""worksheet"":[{""@worksheet_name"":""Sheet1"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Foo""},{""@column_header"":""B"",""#text"":""Bar""},{""@column_header"":""C"",""#text"":""Kanji働""},{""@column_header"":""D"",""#text"":""Summa""}]},{""@row_header"":""2"",""column"":[{""@column_header"":""A"",""#text"":""1""},{""@column_header"":""B"",""#text"":""2""},{""@column_header"":""C"",""#text"":""3""},{""@column_header"":""D"",""#text"":""6""}]}]},{""@worksheet_name"":""OmituinenNimi"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Kissakuva""},{""@column_header"":""B"",""#text"":""1""},{""@column_header"":""C"",""#text"":""2""},{""@column_header"":""D"",""#text"":""3""}]},{""@row_header"":""15"",""column"":{""@column_header"":""A"",""#text"":""Foo""}},{""@row_header"":""16"",""column"":{""@column_header"":""B"",""#text"":""Bar""}}]}]}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }
        [Test]
        public void TestConvertXlsxOneWorksheetToJSON()
        {
            // Test converting one worksheet of xlsx file to JSON
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            var expectedResult = @"{""workbook"":{""@workbook_name"":""ExcelTestInput1.xlsx"",""worksheet"":{""@worksheet_name"":""Sheet1"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Foo""},{""@column_header"":""B"",""#text"":""Bar""},{""@column_header"":""C"",""#text"":""Kanji働""},{""@column_header"":""D"",""#text"":""Summa""}]},{""@row_header"":""2"",""column"":[{""@column_header"":""A"",""#text"":""1""},{""@column_header"":""B"",""#text"":""2""},{""@column_header"":""C"",""#text"":""3""},{""@column_header"":""D"",""#text"":""6""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

        [Test]
        public void TestConvertXlsOneWorksheetToJSON()
        {
            // Test converting one worksheet of xls file to JSON 
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            _options.ReadOnlyWorkSheetWithName = "Sheet1";
            var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
            string expectedResult = @"{""workbook"":{""@workbook_name"":""ExcelTestInput2.xls"",""worksheet"":{""@worksheet_name"":""Sheet1"",""row"":[{""@row_header"":""1"",""column"":[{""@column_header"":""A"",""#text"":""Foo""},{""@column_header"":""B"",""#text"":""Bar""},{""@column_header"":""C"",""#text"":""Kanji働""},{""@column_header"":""D"",""#text"":""Summa""}]},{""@row_header"":""2"",""column"":[{""@column_header"":""A"",""#text"":""1""},{""@column_header"":""B"",""#text"":""2""},{""@column_header"":""C"",""#text"":""3""},{""@column_header"":""D"",""#text"":""6""}]}]}}}";
            Assert.That(Regex.Replace(result.ToJson().ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }


        [Test]
        public void ShouldThrowUnknownFileFormatError()
        {
            // Test converting one worksheet of xls file to csv 
            _input.Path = Path.Combine(_input.Path, "UnitTestErrorFile.txt");
            _options.ThrowErrorOnFailure = true;
            Assert.That(() => ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken()),Throws.Exception);
        }

        [Test]
        public void DoNotThrowOnFailure()
        {
            //try to convert a file that does not exist 
            _input.Path = Path.Combine(_input.Path, "thisfiledoesnotexist.txt");
            _options.ThrowErrorOnFailure = false;
            try
            {
                var result = ExcelTask.ConvertExcelFile(_input, _options, new System.Threading.CancellationToken());
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
