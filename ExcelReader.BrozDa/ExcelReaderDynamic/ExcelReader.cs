using OfficeOpenXml;

using System.Globalization;

using System.Text.RegularExpressions;


namespace ExcelReaderDynamic
{
    internal class ExcelReader
    {
        public string FilePath { get; set; }

        private List<string> _headers = new();

        public ExcelReader(string filePath)
        {
            FilePath = filePath;
        }


        public List<Record> GetAllRecords()
        {
            if (!File.Exists(FilePath))
            {
                return null;
            }
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
        /*private Person MapDataRowToPerson(Dictionary<string, string> dataRow)
        {
            return new Person()
            {
                UserId = dataRow["User Id"],
                FirstName = dataRow["First Name"],
                LastName = dataRow["Last Name"],
                Email = dataRow["Email"],
                Phone = FormatPhoneNumber(dataRow["Phone"]),
                BirthDate = FormatBirthDate(dataRow["Date of birth"])
            };

        }
        private DateTime? FormatBirthDate(string input)
        {
            return DateTime.TryParseExact(input, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var result)
                ? result
                : null;
        }
        private string FormatPhoneNumber(string initialFormat)
        {
            //remove anything but numbers or x, then remove all zeroes from start
            string digitsOnly = Regex.Replace(initialFormat, @"[^0-9x]", "")
                .TrimStart('0');

            return digitsOnly.StartsWith('1')
                ? "+" + digitsOnly
                : "+1" + digitsOnly;

        }*/
    }
}
