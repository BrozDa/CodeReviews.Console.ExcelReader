using ExcelReader.BrozDa.Models;
using OfficeOpenXml;

namespace ExcelReader.BrozDa.Services
{
    internal class ExcelReadingService
    {
        private List<string>? _dataHeaders = null;

        public List<DataModel> ReadExcelFile()
        {
            string path = Path.Combine(Environment.CurrentDirectory, "Resources/people-100.xlsx");

            var data = new List<DataModel>();
            var file = new FileInfo(path);

            using var package = new ExcelPackage(file);

            var excelWorksheet = package.Workbook.Worksheets[0];

            for (int i = 2; i <= excelWorksheet.Dimension.Rows; i++)
            {
                data.Add(GetDataRow(excelWorksheet, i));
            }

            return data;

        }
        private DataModel GetDataRow(ExcelWorksheet excelWorksheet, int row)
        {
            if (_dataHeaders == null)
            {
                _dataHeaders = GetData(excelWorksheet, 1);
            }
            
            return new DataModel() { Headers = _dataHeaders, Data = GetData(excelWorksheet, row) };
            
        }
        private List<string> GetData(ExcelWorksheet excelWorksheet, int row)
        {
            List<string> data = new();
            for (int i = 1; i <= excelWorksheet.Dimension.Columns; i++) 
            {
                data.Add(excelWorksheet.Cells[row, i].Text);

            }

            return data;
        }
    }
}
