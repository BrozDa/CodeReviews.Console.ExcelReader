using OfficeOpenXml;
using CsvHelper;
using System.Globalization;
using Spectre.Console;
using System.Net.WebSockets;

namespace ExcelReaderDynamic.Services
{
    internal class ExcelReaderService
    {

        public List<Record> ReadAllRecords(string filePath)
        {
            using var package = new ExcelPackage(filePath);

            var excelWorksheet = package.Workbook.Worksheets[0];

            List<string> headers = GetHeaders(excelWorksheet);
            List<Record> records = new();

            for (int row = 2; row <= excelWorksheet.Dimension.Rows; row++)
            {
                records.Add(new Record()
                {
                    Headers = headers,
                    Data = GetData(excelWorksheet, row),
                });
            }

            return records;

        }
        private List<string> GetHeaders(ExcelWorksheet sheet)
        {
            List<string> headers = new();

            for (int i = 1; i <= sheet.Dimension.Columns; i++)
            {
                headers.Add(sheet.Cells[1, i].Text);
            }
            return headers;
        }
        
        private List<string> GetData(ExcelWorksheet sheet, int row)
        {
            List<string> data = new();

            for (int col = 1; col <= sheet.Dimension.Columns; col++)
            {
                data.Add(sheet.Cells[row, col].Text);
            }

            return data;
        }

        public void WriteToFile(string filePath, List<Record> records)
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Data");

            WriteHeaders(sheet, records[0].Headers);

            int row = 2;

            foreach (var record in records) 
            {
                WriteData(sheet, record.Data, row);
                row++;
            }
            FileStream objFileStrm = File.Create(filePath);
            objFileStrm.Close();

            File.WriteAllBytes(filePath, package.GetAsByteArray());
            Console.WriteLine("Done more");
            Console.ReadKey();

        }
        public void WriteHeaders(ExcelWorksheet sheet, List<string> headers)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                sheet.Cells[1, i+1].Value = headers[i];
            }
        }
        public void WriteData(ExcelWorksheet sheet, List<string> data, int row)
        {
            for (int i = 0; i < data.Count; i++)
            {
                sheet.Cells[row, i+1].Value = data[i];
            }
        }
    }
}
