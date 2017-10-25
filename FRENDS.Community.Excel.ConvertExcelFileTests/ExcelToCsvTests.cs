using FRENDS.Community;
using FRENDS.Community.Excel.ConvertExcelFile;
using NUnit.Framework;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace FRENDS.Tests
{
    [TestFixture]
    public class ExcelConvertTests
    {
        Input input = new Input();
        Options options = new Options();

        // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required)
        // It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/


        [SetUp]
        public void Setup()
        {
            input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\TestData\ExcelTests\In\");
            options.CsvSeparator = ",";
            options.ReadOnlyWorkSheetWithName = "";
            
        }
        [Test]
        public void TestConvertXlsxToCSV()
        {

            // Test converting all worksheets of xlsx file to csv
            input.Path = Path.Combine(input.Path, "ExcelTestInput1.xlsx");
            options.outputFileType = OutputFileType.CSV;
            var result = ExcelClass.ConvertExcelFile(input, options);
            string expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
        }

       [Test]
       public void TestConvertXlsToCSV()
       {
           // Test converting all worksheets of xls file to csv 
           input.Path = Path.Combine(input.Path, "ExcelTestInput2.xls");
           options.outputFileType = OutputFileType.CSV;
           var result = ExcelClass.ConvertExcelFile(input, options);
           string expectedResult = "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
           Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
       }

       [Test]
       public void TestConvertXlsxToXML()
       {
           // Test converting all worksheets of xlsx file to xml 
           input.Path = Path.Combine(input.Path, "ExcelTestInput1.xlsx");
           options.outputFileType = OutputFileType.XML;
           var result = ExcelClass.ConvertExcelFile(input, options);
           string expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
           Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
       }

       [Test]
       public void TestConvertXlsToXML()
       {
           // Test converting all worksheets of xls file to xml 
           input.Path = Path.Combine(input.Path, "ExcelTestInput2.xls");
           options.outputFileType = OutputFileType.XML;
           var result = ExcelClass.ConvertExcelFile(input, options);
           string expectedResult = @"<workbookworkbook_name=""ExcelTestInput2.xls""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
           Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
       }

       [Test]
       public void TestConvertXlsxOneWorksheetToXML()
       {
           // Test converting one worksheet of xlsx file to xml 
           input.Path = Path.Combine(input.Path, "ExcelTestInput1.xlsx");
           options.ReadOnlyWorkSheetWithName = "Sheet1";
           options.outputFileType = OutputFileType.XML;
           var result = ExcelClass.ConvertExcelFile(input, options);
           string expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet></workbook>";
           Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
       }

       [Test]
       public void TestConvertXlsOneWorksheetToCSV()
       {
           // Test converting one worksheet of xls file to csv 
           input.Path = Path.Combine(input.Path, "ExcelTestInput2.xls");
           options.outputFileType = OutputFileType.CSV;
           var result = ExcelClass.ConvertExcelFile(input, options);
           string expectedResult = "Foo,Bar,Kanji働,Summa1,2,3,6";
           Assert.That(Regex.Replace(result.resultData.ToString(), @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
       }

        [Test]
        public void ShouldThrowUnknownFileFormatError()
        {
            // Test converting one worksheet of xls file to csv 
            input.Path = Path.Combine(input.Path, "UnitTestErrorFile.txt");
            options.outputFileType = OutputFileType.CSV;

            try
            {
                var result = ExcelClass.ConvertExcelFile(input, options);
            }
            catch (ArgumentException ex)
            {
                Assert.AreEqual("Unknown input file type. Please use .xlsx or .xls.", ex.Message);
            }
            catch(Exception e)
            {
                Assert.Fail(string.Format("Unexpected exception of type {0} caught: {1}", e.GetType(), e.Message));
            }
        }

    }
}
