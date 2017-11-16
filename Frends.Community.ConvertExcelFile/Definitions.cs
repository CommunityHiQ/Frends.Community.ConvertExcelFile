using Frends.Tasks.Attributes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.ComponentModel;
using System.Xml;

#pragma warning disable 1591

namespace Frends.Community.Excel.ConvertExcelFile
{
    public class Input
    {
        /// <summary>
        /// Path to the Excel file
        /// </summary>
        [DefaultValue(@"C:\tmp\ExcelFile.xlsx")]
        [DefaultDisplayType(DisplayType.Text)]
        public string Path { get; set; }
    }
    
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
        /// Excel-conversion to JSON
        /// </summary>
        /// <returns>JToken</returns>
        public object ToJson() { return _json; }

        /// <summary>
        /// Excel-conversion to CSV
        /// </summary>
        /// <returns></returns>
        public string ToCsv() { return _csv; }


        private string _csv;
        private object _json;

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
}
