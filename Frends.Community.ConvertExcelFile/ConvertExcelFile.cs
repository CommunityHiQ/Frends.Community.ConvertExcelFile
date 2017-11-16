using System;
using System.Text;
using System.Data;
using System.IO;
using System.Xml;
using ExcelDataReader;
using System.Threading;

#pragma warning disable 1591

namespace Frends.Community.ConvertExcelFile
{
    public class ExcelClass
    {
        /// <summary>
        /// A Frends-task for converting Excel-files to XML, CSV and JSON
        /// </summary>
        /// <returns>Object {string ResultData, bool Success, string Message, JToken ToJson(), string ToCsv()}</returns>
        public static Result ConvertExcelFile(Input input, Options options, CancellationToken cancellationToken)
        {
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

                        return new Result(true, resultDataXML, resultCSV);
                    }
                }
            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnFailure)
                {
                    throw new Exception(ex.ToString());
                }
                return new Result(false, ex.ToString());
            }
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