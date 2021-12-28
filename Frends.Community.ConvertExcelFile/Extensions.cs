using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using Newtonsoft.Json.Linq;

namespace Frends.Community.ConvertExcelFile
{
    class Extensions
    {
        /// <summary>
        /// Converts DataSet-object to JSON.
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="file_name">Excel file name to be read</param>
        /// <returns>JObject containing the converted Excel.</returns>
        public static object WriteJToken(DataSet result, Options options, string file_name)
        {
            var json = ConvertToJson(result, options, file_name);
            return JObject.Parse(json.ToString());
        }
        /// <summary>
        /// Converts DataSet-object to JSON.
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="file_name">Excel file name to be read</param>
        /// <returns>JToken containing the converted Excel.</returns>
        public static object ConvertToJson(DataSet result, Options options, string file_name)
        {
            var json = new StringBuilder();
            json.Append("{");
            json.Append($"\"workbook\": ");
            json.Append("{");
            json.Append($"\"workbook_name\": \"{file_name}\",");
            if (options.ReadOnlyWorkSheetWithName.Length == 0)
            {
                json.Append("\"worksheets\": ");
                json.Append("[");
            }
            else
            {
                json.Append("\"worksheet\" : ");
            }

            foreach (DataTable dt in result.Tables)
            {

                if (options.ReadOnlyWorkSheetWithName.Contains(dt.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                {
                    json.Append("{");
                    json.Append($"\"name\": \"{dt.TableName}\",");
                    json.Append("\"rows\": ");

                    // building json from datatable
                    if (dt.Rows.Count > 0)
                    {
                        json.Append("[");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            var content = WriteRowToJson(dt, i, options).ToString();
                            if (!content.ToString().Equals("empty"))
                            {
                                json.Append(WriteRowToJson(dt, i, options).ToString());
                                if (i < dt.Rows.Count - 1)
                                {
                                    json.Append("},");
                                }
                                else if (i == dt.Rows.Count - 1)
                                {
                                    json.Append("}");
                                }
                            }
                        }

                        json.Append("]");
                    }
                    if (result.Tables.IndexOf(dt) != result.Tables.Count - 1)
                    {
                        json.Append("},");
                    }
                    else
                    {
                        json.Append("}");
                    }
                }

            }
            if (options.ReadOnlyWorkSheetWithName.Length == 0)
            {
                json.Append("]");
            }
            json.Append("}");
            json.Append("}");

            return json;

        }
        /// <summary>
        /// Converts DataRow-object to JSON as a StringBuilder object.
        /// </summary>
        /// <param name="dt">DataTable-object</param>
        /// <param name="i">Iteration index</param>
        /// <param name="options">Input configurations</param>
        /// <returns>StringBuilder-object containing the converted Excel from row.</returns>
        public static object WriteRowToJson(DataTable dt, int i, Options options)
        {
            var rowJson = new StringBuilder();
            rowJson.Append("{");
            rowJson.Append($"\"{i + 1}\":");

            var content = WriteColumnToJson(dt, i, options).ToString();
            if (content.Equals("[]"))
            {
                return "empty";
            }

            rowJson.Append(content);

            return rowJson;
        }
        /// <summary>
        /// Converts DataColumn-object to JSON as StringBuilder object.
        /// </summary>
        /// <param name="dt">DataTable-object</param>
        /// <param name="i">Iteration index</param>
        /// <param name="options">Input configurations</param>
        /// <returns>StringBuilder-object containing the converted Excel from column.</returns>
        public static object WriteColumnToJson(DataTable dt, int i, Options options)
        {
            var columnJson = new StringBuilder();
            columnJson.Append("[");
            for (var j = 0; j < dt.Columns.Count; j++)
            {
                var content = dt.Rows[i].ItemArray[j];
                var type = content.GetType();
                if (string.IsNullOrWhiteSpace(content.ToString()) == false)
                {
                    if (content.GetType().Name == "DateTime")
                    {
                        content = ConvertDateTimes((DateTime)content, options);
                    }
                    content = content.ToString();
                    columnJson.Append("{");
                    if (options.UseNumbersAsColumnHeaders)
                    {
                        columnJson.Append($"\"{j + 1}\":");
                    }
                    else
                    {
                        columnJson.Append($"\"{ColumnIndexToColumnLetter(j + 1)}\":");
                    }
                    columnJson.Append($"\"{content}\"");
                    if (j != dt.Columns.Count - 1)
                    {
                        columnJson.Append("},");
                    }
                    else
                    {
                        columnJson.Append("}");
                    }

                }

            }
            columnJson.Append("]");

            return columnJson;
        }
        /// <summary>
        /// Converts DataSet-object to XML.
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="file_name">Excel file name to be read</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing contents in XML format.</returns>
        public static string ConvertToXml(DataSet result, Options options, string file_name, CancellationToken cancellationToken)
        {
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
                    return builder.ToString();
                }
            }
        }
        /// <summary>
        /// Converts DataSet-object to CSV.
        /// </summary>
        /// <param name="result">DataSet-object</param>
        /// <param name="options">Input configurations</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>String containing the converted Excel.</returns>
        public static string ConvertToCSV(DataSet result, Options options, CancellationToken cancellationToken)
        {
            var resultData = new StringBuilder();

            foreach (DataTable table in result.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Read only wanted worksheets. If none is specified read all. //
                if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
                {
                    for (var i = 0; i < table.Rows.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        for (var j = 0; j < table.Columns.Count; j++)
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
        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            int mod;
            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = ((div - mod) / 26);
            }
            return colLetter;
        }

        /// <summary>
        /// a Helper method 
        /// Converts DateTime object to the DateFormat given as options
        /// Return agent's date format in default
        /// </summary>
        /// <param name="date"></param>
        /// <param name="options"></param>
        /// <returns>string containing correct date format</returns>
        public static string ConvertDateTimes(DateTime date, Options options)
        {
            // modify the date using date format var in options
            
            if (options.ShortDatePattern) 
            {
                switch (options.DateFormat)
                {
                    case DateFormats.DDMMYYYY:
                        return date.ToString(new CultureInfo("en-FI").DateTimeFormat.ShortDatePattern);
                    case DateFormats.MMDDYYYY:
                        return date.ToString(new CultureInfo("en-US").DateTimeFormat.ShortDatePattern);
                    case DateFormats.YYYYMMDD:
                        return date.ToString(new CultureInfo("ja-JP").DateTimeFormat.ShortDatePattern);
                    case DateFormats.DEFAULT:
                        return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                    default:
                        return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                }
            }
            else
            {
                switch (options.DateFormat)
                {
                    case DateFormats.DDMMYYYY:
                        return date.ToString(new CultureInfo("en-FI"));
                    case DateFormats.MMDDYYYY:
                        return date.ToString(new CultureInfo("en-US"));
                    case DateFormats.YYYYMMDD:
                        return date.ToString(new CultureInfo("ja-JP"));
                    case DateFormats.DEFAULT:
                        return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat);
                    default:
                        return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat);
                }
            }
            
        }
    }
}
