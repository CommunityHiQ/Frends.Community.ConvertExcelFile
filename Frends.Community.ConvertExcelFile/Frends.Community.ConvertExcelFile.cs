using System;
using System.IO;
using System.Text;
using System.Threading;
using ExcelDataReader;

#pragma warning disable 1591

namespace Frends.Community.ConvertExcelFile
{

    public class ExcelClass
    {
        /// <summary>
        /// A Frends-task for converting Excel-files to DataSet, XML, CSV and JSON.
        /// </summary>
        /// <returns>Object {DataSet ResultData, bool Success, string Message, JToken ToJson(), string ToXml(), string ToCsv()}</returns>
        public static Result ConvertExcelFile(Input input, Options options, CancellationToken cancellationToken)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using (var stream = new FileStream(input.Path, FileMode.Open))
                {
                    using (var excelReader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = excelReader.AsDataSet();
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
