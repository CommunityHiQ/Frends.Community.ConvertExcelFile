using System;
using System.Text;
using Excel;
using System.Data;
using System.IO;
using System.ComponentModel;
using System.Xml;
using Frends.Tasks.Attributes;

namespace FRENDS.Community.Excel.ConvertExcelFile
{
    /// <summary>
    /// Input class
    /// </summary>
    public class Input
    {
        /// <summary>
        /// Path of the Excel file to be read.
        /// </summary>
        [DefaultValue(@"C:\tmp\ExcelFile.xlsx")]
        [DefaultDisplayType(DisplayType.Text)]
        public string Path { get; set; }

    }

    /// <summary>
    /// Format output string as XML or CSV.
    /// </summary>
    public enum OutputFileType
    {
        /// <summary>
        /// Format output string as XML.
        /// </summary>
        XML,

        /// <summary>
        /// Format output string as CSV.
        /// </summary>
        CSV
    }

    /// <summary>
    /// Options class
    /// </summary>
    public class Options
    {
        /// <summary>
        /// If empty, all work sheets are read.
        /// </summary>
        [DefaultValue(@"")]
        public string ReadOnlyWorkSheetWithName { get; set; }

        /// <summary>
        /// Format output string as XML or CSV.
        /// </summary>
        public OutputFileType outputFileType { get; set; }

        /// <summary>
        /// Csv separator
        /// </summary>
        [ConditionalDisplay(nameof(outputFileType), OutputFileType.CSV)]
        [DefaultValue(@";")]
        [DefaultDisplayType(DisplayType.Text)]
        public string CsvSeparator { get; set; }

        /// <summary>
        /// If set to true, outputs column headers as numbers instead of letters.
        /// </summary>
        [DefaultValue(false)]
        [ConditionalDisplay(nameof(outputFileType), OutputFileType.XML)]
        public bool UseNumbersAsColumnHeaders { get; set; }
    }

    /// <summary>
    /// Result class
    /// </summary>
    public class Result
    {
        /// <summary>
        /// Outputs converted csv/xml data
        /// </summary>
        public string resultData { get; set; }
    }

    /// <summary>
    /// ExcelClass
    /// </summary>
    public class ExcelClass
    {
        /// <summary>
        /// Reads Excel files and converts it to XML or CSV according to the task input parameters.
        /// </summary>
        /// <returns>String containing the file contents as XML or CSV format</returns>
        public static Result ConvertExcelFile(Input input, Options options)
        {
            Result resultData = new Result();
            FileStream stream = new FileStream(input.Path, FileMode.Open);
            IExcelDataReader excelReader;
            var input_filetype = Path.GetExtension(input.Path).ToLower();

            if (input_filetype == ".xlsx")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else if (input_filetype == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                throw new ArgumentException("Unknown input file type. Please use .xlsx or .xls.");
            }

            DataSet result = excelReader.AsDataSet();
          

            if (options.outputFileType == OutputFileType.XML)
            {
                resultData = ConvertToXml(excelReader, result, options, Path.GetFileName(input.Path));
            }
            else
            {
                resultData = ConvertToCSV(excelReader, result, options);
            }
            excelReader.Close();
            return resultData;
        }

        /// <summary>
        /// Converts column header index to letter, as Excel does in its GUI.
        /// </summary>
        /// <returns>String containing correct letter combination for column.</returns>
        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;
            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        /// <summary>
        /// Converts IExcelDataReader object to XML.
        /// </summary>
        /// <returns>String containing contents in XML format</returns>
        private static Result ConvertToXml(IExcelDataReader excelReader, DataSet result, Options options, string file_name)
        {
            Result resultClass = new Result();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;

            StringBuilder builder = new StringBuilder();
            using (StringWriter sw = new StringWriter(builder))
            {
                using (XmlWriter xw = XmlWriter.Create(sw, settings))
                {
                    // Write workbook element. Workbook is also known as sheet.
                    xw.WriteStartDocument();
                    xw.WriteStartElement("workbook");
                    xw.WriteAttributeString("workbook_name", file_name);

                    foreach (DataTable table in result.Tables)
                    {
                        // Read only wanted worksheets. If none is specified read all.
                        if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                        {
                            // Write worksheet element
                            xw.WriteStartElement("worksheet");
                            xw.WriteAttributeString("worksheet_name", table.TableName);

                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                bool row_element_is_writed = false;
                                for (int j = 0; j < table.Columns.Count; j++)
                                {
                                    // Write column only if it has some content
                                    string content = table.Rows[i].ItemArray[j].ToString();
                                    if (String.IsNullOrWhiteSpace(content) == false)
                                    {

                                        if (row_element_is_writed == false)
                                        {
                                            xw.WriteStartElement("row");
                                            xw.WriteAttributeString("row_header", (i + 1).ToString());
                                            row_element_is_writed = true;
                                        }

                                        xw.WriteStartElement("column");
                                        if (options.UseNumbersAsColumnHeaders) {
                                            xw.WriteAttributeString("column_header", (j + 1).ToString());
                                        }
                                        else
                                        {
                                            xw.WriteAttributeString("column_header", ColumnIndexToColumnLetter(j + 1));
                                        }
                                        xw.WriteString(content);
                                        xw.WriteEndElement();
                                    }
                                }
                                if (row_element_is_writed == true)
                                {
                                    xw.WriteEndElement();
                                }
                            }
                            xw.WriteEndElement();
                        }
                    }
                    xw.WriteEndDocument();
                    xw.Close();
                    resultClass.resultData += builder.ToString();
                }
                return resultClass;
            }
        }

        /// <summary>
        /// Converts IExcelDataReader object to CSV.
        /// </summary>
        /// <returns>String containing contents in CSV format</returns>
        private static Result ConvertToCSV(IExcelDataReader excelReader, DataSet result, Options options)
        {
            Result resultClass = new Result();
            foreach (DataTable table in result.Tables)
            {
                // Read only wanted worksheets. If none is specified read all. //
                if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            resultClass.resultData += table.Rows[i].ItemArray[j];
                            if (j < table.Columns.Count - 1)
                            {
                                resultClass.resultData += options.CsvSeparator;
                            }
                        }
                        resultClass.resultData += "\n";
                    }
                }
            }
            return resultClass;

        }

    }
}


