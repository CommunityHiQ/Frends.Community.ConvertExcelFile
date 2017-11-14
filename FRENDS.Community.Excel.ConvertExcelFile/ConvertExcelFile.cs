using System;
using System.Text;
using System.Data;
using System.IO;
using System.ComponentModel;
using System.Xml;
using Frends.Tasks.Attributes;
using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading;

namespace FRENDS.Community.Excel.ConvertExcelFile
{
    /// <summary>
    /// ExcelClass
    /// </summary>
    public class ExcelClass
    {
        /// <summary>
        /// Input class
        /// </summary>
        public class Input
        {
        /// <summary>
        /// Path to the Excel file
        /// </summary>
        [DefaultValue(@"C:\tmp\ExcelFile.xlsx")]
        [DefaultDisplayType(DisplayType.Text)]
        public string Path { get; set; }
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
            /// Csv separator
            /// </summary>
            [DefaultValue(@";")]
            [DefaultDisplayType(DisplayType.Text)]
            public string CsvSeparator { get; set; }
            /// <summary>
            /// If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...) 
            /// </summary>
            [DefaultValue("false")]
            public bool UseNumbersAsColumnHeaders { get; set; }
            /// <summary>
            /// Choose if exception should be thrown when conversion fails.
            /// </summary>
            [DefaultValue("true")]
            public bool ThrowErrorOnFailure { get; set; }
        }
        /// <summary>
        /// Result class
        /// </summary>
        public class Result
        {
            /// <summary>
            /// Converted Excel in XML-format
            /// </summary>
            [DefaultValue("")]
            public string ResultData { get; set; }
            /// <summary>
            /// False if conversion fails
            /// </summary>
            [DefaultValue("false")]
            public Boolean Success { get; set; }
            /// <summary>
            /// Exception message
            /// </summary>
            [DefaultValue("")]
            public string Message { get; set; }
            /// <summary>
            /// Converted Excel in CSV-format
            /// </summary>
            private string _csv;
            /// <summary>
            /// Converted XML in JToken
            /// </summary>
            private object _json;
            /// <summary>
            /// Excel-conversion to JSON
            /// </summary>
            /// <returns>JToken</returns>
            public object ToJson(){return _json;}
            /// <summary>
            /// Excel-conversion to CSV
            /// </summary>
            /// <returns></returns>
            public string ToCsv(){return _csv;}
            /// <summary>
            /// Constructor for successful conversion
            /// </summary>
            /// <param name="success">true if conversion was successful</param>
            /// <param name="resultData">converted Excel in XML-format</param>
            /// <param name="csv">converted Excel in CSV-format</param>
            public Result(bool success, string resultData, string csv)
            {
                Success = success;
                ResultData = resultData;
                _csv = csv;

                if (resultData != null)
                {
                    //creates a JToken from XML
                    var doc = new XmlDocument();
                    doc.LoadXml(resultData);
                    var jsonString = JsonConvert.SerializeXmlNode(doc);
                    _json = JToken.Parse(jsonString);
                }
            }
            /// <summary>
            /// constructor for failed conversion
            /// </summary>
            /// <param name="success">false if conversion failed</param>
            /// <param name="message">holds the exception message</param>
            public Result(bool success, string message)
            {
                Success = success;
                Message = message;
            }
        }

        /// <summary>
        /// A Frends-task for converting Excel-files to XML, CSV and JSON
        /// </summary>
        /// <returns>Object {string ResultData, bool Success, string Message, JToken ToJson(), string ToCsv()}</returns>
        public static Result ConvertExcelFile(Input input, Options options, CancellationToken cancellationToken)
        {
            Result resultData;
            try
            {
                using (FileStream stream = new FileStream(input.Path, FileMode.Open))
                {
                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var input_filetype = Path.GetExtension(input.Path).ToLower();
                        DataSet result = excelReader.AsDataSet();
                        //Convert Excel to XML and CSV
                        string resultDataXML = ConvertToXml(excelReader, result, options, Path.GetFileName(input.Path), cancellationToken);
                        string resultCSV = ConvertToCSV(result, options, cancellationToken);
                        resultData = new Result(true, resultDataXML, resultCSV);
                    }
                }
            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnFailure)
                {
                    throw new Exception(ex.ToString());
                }
                resultData = new Result(false, ex.ToString());
            }
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
        /// <param name="excelReader">Interface for  reading Excel data.</param>
        /// <param name="result">Excel DataSet</param>
        /// <param name="options">Input configurations</param>
        /// <param name="file_name">Excel file name to be read</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing contents in XML format</returns>
        private static string ConvertToXml(IExcelDataReader excelReader, DataSet result, Options options, string file_name, CancellationToken cancellationToken)
        {
            String xml_string;

            XmlWriterSettings settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true
            };

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
                        cancellationToken.ThrowIfCancellationRequested();
                        // Read only wanted worksheets. If none is specified read all.
                        if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                        {
                            // Write worksheet element
                            xw.WriteStartElement("worksheet");
                            xw.WriteAttributeString("worksheet_name", table.TableName);

                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                bool row_element_is_writed = false;
                                for (int j = 0; j < table.Columns.Count; j++)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
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
                                        if (options.UseNumbersAsColumnHeaders)
                                        {
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
                    xml_string = builder.ToString();
                }
                return xml_string;
            }
        }
        /// <summary>
        /// Converts IExcelDataReader object to CSV
        /// </summary>
        /// <param name="result">Excel DataSet</param>
        /// <param name="options">Input configurations</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing the converted Excel</returns>
        private static string ConvertToCSV(DataSet result, Options options, CancellationToken cancellationToken)
        {
            string resultData = null;

            foreach (DataTable table in result.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Read only wanted worksheets. If none is specified read all. //
                if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            resultData += table.Rows[i].ItemArray[j];
                            if (j < table.Columns.Count - 1)
                            {
                                resultData += options.CsvSeparator;
                            }
                        }
                        resultData += "\n";
                    }
                }
            }
            return resultData;
        }
    }
}