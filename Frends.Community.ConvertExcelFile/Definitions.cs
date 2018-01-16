using Frends.Tasks.Attributes;
using System;
using System.ComponentModel;
using System.Data;
using System.Threading;

#pragma warning disable 1591

namespace Frends.Community.ConvertExcelFile
{
    public class Input
    {
        /// <summary>
        /// Path to the Excel file.
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
        /// Csv separator.
        /// </summary>
        [DefaultValue(@";")]
        [DefaultDisplayType(DisplayType.Text)]
        public string CsvSeparator { get; set; }

        /// <summary>
        /// If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...). 
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
        /// Converted Excel in DataSet-format.
        /// </summary>
        [DefaultValue(null)]
        public DataSet ResultData { get; set; }
        /// <summary>
        /// False if conversion fails.
        /// </summary>
        [DefaultValue("false")]
        public Boolean Success { get; set; }
        /// <summary>
        /// Exception message.
        /// </summary>
        [DefaultValue("")]
        public string Message { get; set; }
        /// <summary>
        /// Excel-conversion to JSON.
        /// </summary>
        /// <returns>JToken</returns>
        public object ToJson() { return _json.Value; }
        /// <summary>
        /// Excel-conversion to CSV.
        /// </summary>
        /// <returns>String</returns>
        public string ToCsv() { return _csv.Value; }
        /// <summary>
        /// Excel-conversion to XML.
        /// </summary>
        /// <returns>String</returns>
        public string ToXml() { return _xml.Value; }
        private readonly Lazy<string> _csv;
        private readonly Lazy<object> _json;
        private readonly Lazy<string> _xml;
        //Constructor for successful conversion
        public Result(bool success, DataSet result, Options options, string filename, CancellationToken cancellationToken)
        {
            Success = success;
            ResultData = result;

            _xml = new Lazy<string>(() => ResultData != null ? HelperMethods.ConvertToXml(ResultData,  options, filename, cancellationToken) : null);
            _json = new Lazy<object>(() => ResultData != null ? HelperMethods.WriteJToken(ResultData, options, filename,cancellationToken) : null);
            _csv = new Lazy<string>(() => ResultData != null ? HelperMethods.ConvertToCSV(ResultData, options, cancellationToken) : null);
        }
        //Constructor for failed conversion
        public Result(bool success, string message)
        {
            Success = success;
            Message = message;
            ResultData = null;
        }
    }    
}
