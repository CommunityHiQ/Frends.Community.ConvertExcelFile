using System;
using System.Text;
using System.Data;
using System.IO;
using System.ComponentModel;
using System.Xml;
using Frends.Tasks.Attributes;
using ExcelDataReader;
using Newtonsoft.Json.Linq;

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
        CSV,
        /// <summary>
        /// Format output string as JSON
        /// </summary>
        JSON
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
        public OutputFileType OutputFileType { get; set; }

        /// <summary>
        /// Csv separator
        /// </summary>
        [ConditionalDisplay(nameof(OutputFileType), OutputFileType.CSV)]
        [DefaultValue(@";")]
        [DefaultDisplayType(DisplayType.Text)]
        public string CsvSeparator { get; set; }

        /// <summary>
        /// If set to true, outputs column headers as numbers instead of letters.
        /// </summary>
        [DefaultValue(false)]
        [ConditionalDisplay(nameof(OutputFileType), OutputFileType.XML)]
        public bool UseNumbersAsColumnHeaders { get; set; }
        /// <summary>
        /// Choose if exception should be thrown when conversion fails.
        /// </summary>
        [DefaultValue(true)]
        public bool ThrowErrorOnfailure { get; set; }
    }
    /// <summary>
    /// Result class
    /// </summary>
    public class Result
    {
        /// <summary>
        /// Outputs converted csv/xml data
        /// </summary>
        [DefaultValue("")]
        public string ResultData { get; set; }
        /// <summary>
        /// False if conversion fails
        /// </summary>
        [DefaultValue(false)]
        public Boolean Success { get; set; }
        /// <summary>
        /// Exception message
        /// </summary>
        [DefaultValue("")]
        public string Message { get; set; }
    }
    /// <summary>
    /// ExcelClass
    /// </summary>
    public class ExcelClass
    {
        /// <summary>
        /// Reads Excel files and converts it to XML JSON or CSV according to the task input parameters.
        /// </summary>
        /// <returns>String containing the file contents as XML or CSV format</returns>
        public static Result ConvertExcelFile(Input input, Options options)
        {
            Result resultData = new Result
            {
                Success = false
            };

            try
            {
                using (FileStream stream = new FileStream(input.Path, FileMode.Open))
                {
                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var input_filetype = Path.GetExtension(input.Path).ToLower();
                        DataSet result = excelReader.AsDataSet();

                        switch (options.OutputFileType)
                        {
                            case OutputFileType.XML:
                                resultData = ConvertToXml(excelReader, result, options, Path.GetFileName(input.Path));
                                break;
                            case OutputFileType.CSV:
                                resultData = ConvertToCSV(result, options);
                                break;
                            case OutputFileType.JSON:
                                resultData = ConvertToJSON(result, options, Path.GetFileName(input.Path));
                                break;
                        }
                    }
                }
                resultData.Success = true;
            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnfailure)
                {
                    throw new Exception(ex.ToString());
                }
                resultData.Message = ex.ToString();
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
        /// <returns>String containing contents in XML format</returns>
        private static Result ConvertToXml(IExcelDataReader excelReader, DataSet result, Options options, string file_name)
        {
            Result resultClass = new Result();

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
                    resultClass.ResultData += builder.ToString();
                }
                return resultClass;
            }
        }
        /// <summary>
        /// Converts IExcelDataReader object to CSV
        /// </summary>
        /// <param name="result">Excel DataSet</param>
        /// <param name="options">Input configurations</param>
        /// <returns>String containing contents in CSV format</returns>
        private static Result ConvertToCSV(DataSet result, Options options)
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
                            resultClass.ResultData += table.Rows[i].ItemArray[j];
                            if (j < table.Columns.Count - 1)
                            {
                                resultClass.ResultData += options.CsvSeparator;
                            }
                        }
                        resultClass.ResultData += "\n";
                    }
                }
            }
            return resultClass;

        }
        /// <summary>
        /// Converts ExcelReadr DataSet to JSON
        /// </summary>
        /// <returns>Result-object</returns>
        private static Result ConvertToJSON(DataSet result, Options options, string filename)
        {
            Result resultClass = new Result();
            //this will hold the final product 
            JObject workbook = new JObject();
            //container for worksheets
            JArray wsItemContainer = new JArray();
            foreach (DataTable table in result.Tables)
            {
                if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                {
                    //container for all rows
                    JArray rowItemContainer = new JArray();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        bool rowElementIsWritten = false;
                        //a single row
                        JObject rItem = new JObject();
                        //a container for all columns
                        JArray columnItemContainer = new JArray();

                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            string content = table.Rows[i].ItemArray[j].ToString();
                            if (!string.IsNullOrEmpty(content))
                            {
                                //create a new column item with a header(A,B,C..) and a value 
                                JObject cItem = new JObject
                                {
                                    new JProperty("column_header", ColumnIndexToColumnLetter(j+1)),
                                    new JProperty("value", content)
                                };
                                //write the row header/number(1,2,3..) to a row item 
                                if (!rowElementIsWritten)
                                {
                                    rItem.Add(new JProperty("row_header", (i + 1)));
                                    rowElementIsWritten = true;
                                }
                                //add column item to a container
                                columnItemContainer.Add(cItem);
                            }
                        }
                        //if row header is written, add all the columns in the container to the row item
                        //This also ensures that no empty elements are written
                        if (rowElementIsWritten)
                        {
                            rItem.Add(new JProperty("column", columnItemContainer));
                            rowItemContainer.Add(rItem);
                        }
                    }
                    //create a new worksheet, add all rowitems
                    JObject wsItem = new JObject
                    {
                        new JProperty("worksheet_name", table.TableName),
                        new JProperty("row", rowItemContainer)
                    };
                    //worksheet is added to a container holding other worksheets
                    wsItemContainer.Add(wsItem);
                }
                //create a new workbook and add all worksheets
                workbook = new JObject
            {
                new JProperty("workbook_name", filename),
                new JProperty("worksheet", wsItemContainer)
            };

            }
            resultClass.ResultData = workbook.ToString();
            return resultClass;
        }
    }
}