using OfficeOpenXml;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Reflection.PortableExecutable;

namespace ExcelReaderDynamic.Services
{
    internal class ExcelReaderService
    {
        //public string FilePath { get; set; } = string.Empty;
        public string FilePath = @"E:\Git Repos\CodeReviews.Console.ExcelReader\ExcelReader.BrozDa\ExcelReaderDynamic\customers-100.csv";

        private List<string> _headers = new();


        public List<Record> ReadFile()
        {
            if (FilePath.EndsWith(".xlsx"))
            {
                return GetAllRecordsFromXlsx();
            }
            else if (FilePath.EndsWith(".csv"))
            {
                return GetAllRecordsFromCsv();
            }
            else
            {
                Console.WriteLine("Unsuported file type");
                return new List<Record> { new Record() };
            }
        }
        private List<Record> GetAllRecordsFromCsv()
        {
            using StreamReader reader = new StreamReader(FilePath);

            using var csvreader = new CsvReader(reader, CultureInfo.InvariantCulture){};

            csvreader.Read();
            csvreader.ReadHeader();
            _headers = csvreader.HeaderRecord.ToList() ?? new List<string>();

            var records = new List<Record>();   
            while (csvreader.Read()) 
            {
                var data = new List<string>();

                foreach (var header in _headers)
                {
                    data.Add(csvreader.GetField(header));
                }

                records.Add(new Record
                {
                    Headers = _headers,
                    Data = data
                });
            }

            return records;

            

          

            /*var lines = File.ReadAllLines(FilePath);

            _headers = lines[0].Split(',').ToList();

            List<Record> records = lines
                .Skip(1)
                .Select(line => new Record()
                {
                    Headers = _headers,
                    Data = line.Split(',').ToList()
                }).ToList();

            return records;*/
        }
        private List<Record> GetAllRecordsFromXlsx()
        {
            using var package = new ExcelPackage(FilePath);

            var excelWorksheet = package.Workbook.Worksheets[0];

            _headers = GetHeaders(excelWorksheet);
            List<Record> records = new();

            for (int row = 2; row <= excelWorksheet.Dimension.Rows; row++)
            {
                records.Add(new Record()
                {
                    Headers = _headers,
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
    }
}
