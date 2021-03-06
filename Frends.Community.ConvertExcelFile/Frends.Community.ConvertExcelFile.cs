﻿using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using ExcelDataReader;
using Microsoft.CSharp; // You can remove this if you don't need dynamic type in .Net Standard tasks
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

#pragma warning disable 1591

namespace Frends.Community.ConvertExcelFile
{

        public class ExcelClass
        {
            /// <summary>
            /// A Frends-task for converting Excel-files to DataSet, XML, CSV and JSON
            /// </summary>
            /// <returns>Object {DataSet ResultData, bool Success, string Message, JToken ToJson(), string ToXml(), string ToCsv()}</returns>
            public static Result ConvertExcelFile(Input input, Options options, CancellationToken cancellationToken)
            {
                try
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using (FileStream stream = new FileStream(input.Path, FileMode.Open))
                    {
                        using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = excelReader.AsDataSet();
                            return new Result(true, result, options, Path.GetFileName(input.Path), cancellationToken);
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
        }
    

    class HelperMethods
    {
        /// <summary>
        /// Converts DataSet-object to JSON
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="filename">Excel file name to be read</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>a JToken containing the converted Excel</returns>
        internal static object WriteJToken(DataSet result, Options options, string filename, CancellationToken cancellationToken)
        {
            var doc = new XmlDocument();
            doc.LoadXml(ConvertToXml(result, options, filename, cancellationToken));
            var jsonString = JsonConvert.SerializeXmlNode(doc);
            return JToken.Parse(jsonString);
        }
        /// <summary>
        /// Converts DataSet-object to XML.
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="file_name">Excel file name to be read</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing contents in XML format</returns>
        internal static string ConvertToXml(DataSet result, Options options, string file_name, CancellationToken cancellationToken)
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
        /// Converts DataSet-object to CSV
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing the converted Excel</returns>
        internal static string ConvertToCSV(DataSet result, Options options, CancellationToken cancellationToken)
        {
            var resultData = new StringBuilder();

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
                            resultData.Append(table.Rows[i].ItemArray[j] + options.CsvSeparator);
                        }
                        // remove last CsvSeparator
                        resultData.Length--;
                        resultData.Append(Environment.NewLine);
                    }
                }
            }
            return resultData.ToString();
        }

        /// <summary>
        /// a Helper method
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
    }
}
