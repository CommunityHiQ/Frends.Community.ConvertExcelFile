﻿using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
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
        [DisplayFormat(DataFormatString = "Text")]
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
        [DisplayFormat(DataFormatString = "Text")]
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

        /// <summary>
        /// Date format selection.
        /// </summary>
        [DisplayName("Date Format")]
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue(DateFormats.DEFAULT)]
        public DateFormats DateFormat {  get; set; }

        /// <summary>
        /// If set to true, dates will exclude timestamps from dates.
        /// Default false
        /// </summary>
        [DefaultValue("false")]
        public bool ShortDatePattern { get; set; }
    }

    #region Enumerations
    public enum DateFormats
    {
        /// <summary>
        /// default value
        /// </summary>
        DEFAULT,
        /// <summary>
        /// day/month/year
        /// </summary>
        DDMMYYYY,
        /// <summary>
        /// month/day/year
        /// </summary>
        MMDDYYYY,
        /// <summary>
        /// year/month/day
        /// </summary>
        YYYYMMDD
    }
    #endregion

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
        public bool Success { get; set; }
        /// <summary>
        /// Exception message.
        /// </summary>
        [DefaultValue("")]
        public string Message { get; set; }
        /// <summary>
        /// Excel-conversion to JSON.
        /// </summary>
        /// <returns>JObject</returns>
        public object ToJson(){ return _json.Value;}
        /// <summary>
        /// Excel-conversion to CSV.
        /// </summary>
        /// <returns>String</returns>
        public string ToCsv() { return _csv.Value;}
        /// <summary>
        /// Excel-conversion to XML.
        /// </summary>
        /// <returns>String</returns>
        public string ToXml() { return _xml.Value; }
        private readonly Lazy<string> _csv;
        private readonly Lazy<object> _json;
        private readonly Lazy<string> _xml;
        // Constructor for successful conversion.
        public Result(bool success, DataSet result, Options options, string filename, CancellationToken cancellationToken)
        {
            Success = success;
            ResultData = result;

            _xml = new Lazy<string>(() => ResultData != null ? Extensions.ConvertToXml(ResultData,  options, filename, cancellationToken) : null);
            _json = new Lazy<object>(() => ResultData != null ? Extensions.WriteJToken(ResultData, options, filename) : null);
            _csv = new Lazy<string>(() => ResultData != null ? Extensions.ConvertToCSV(ResultData, options, cancellationToken) : null);
        }
        // Constructor for failed conversion.
        public Result(bool success, string message)
        {
            Success = success;
            Message = message;
            ResultData = null;
        }
    }    
}
