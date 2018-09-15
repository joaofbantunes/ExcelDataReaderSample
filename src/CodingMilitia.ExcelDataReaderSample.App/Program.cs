using System;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace CodingMilitia.ExcelDataReaderSample.App
{
    class Program
    {
        static void Main(string[] args)
        {
            // required because of known bug when running on .NET Core
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open("Files/sample.xlsx", FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                DataTableCollection sheets = reader.AsDataSet(GetDataSetConfig()).Tables;
                DataTable sheet = sheets["Sheet1"];
                foreach (DataRow row in sheet.Rows)
                {
                    Console.WriteLine("{0}: {1}", row["Id"], row["Description"]);
                }

            }
        }

        private static ExcelDataSetConfiguration GetDataSetConfig()
        {
            return new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    ReadHeaderRow = rowReader => Console.WriteLine("{0}: {1}", rowReader[0], rowReader[1])
                }
            };
        }
    }
}
