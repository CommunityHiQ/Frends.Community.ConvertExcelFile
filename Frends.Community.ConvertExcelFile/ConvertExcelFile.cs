using System;
using System.Data;
using System.IO;
using ExcelDataReader;
using System.Threading;

#pragma warning disable 1591

namespace Frends.Community.ConvertExcelFile
{
    public class ExcelTask
    {
        /// <summary>
        /// A Frends-task for converting Excel-files to DataSet, XML, CSV and JSON
        /// </summary>
        /// <returns>Object {DataSet ResultData, bool Success, string Message, JToken ToJson(), string ToXml(), string ToCsv()}</returns>
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
}