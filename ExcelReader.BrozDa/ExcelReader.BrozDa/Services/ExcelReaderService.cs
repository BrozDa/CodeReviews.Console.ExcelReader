using ExcelReader.Brozda.Models;
using OfficeOpenXml;

namespace ExcelReaderDynamic.Services
{
    /// <summary>
    /// Handles reading from and writing to XLSX file
    /// </summary>
    internal class ExcelReaderService
    {
        /// <summary>
        /// Reads all records from XLSX file
        /// </summary>
        /// <param name="filePath">A <see cref="string"/> representing path to the file</param>
        /// <returns>A <see cref="ReadingResult{T}"/> indicating wheter the operation was successful
        /// A <see cref="ReadingResult{T}"/> contains retrieved data in form of List of <see cref="Record"/> in case of success
        /// or contains error message in case of any error</returns>
        public ReadingResult<List<Record>> ReadAllRecords(string filePath)
        {
            try
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
                        Data = GetDataRow(excelWorksheet, row),
                    });
                }

                return ReadingResult<List<Record>>.Success(records);
            }
            catch (Exception ex)
            {
                return ReadingResult<List<Record>>.Fail($"Error while reading the file: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves column headers from XLSX file
        /// </summary>
        /// <param name="sheet">A <see cref="ExcelWorksheet"/> from which data will be read</param>
        /// <returns>A list of <see cref="string"/> containing headers from file</returns>
        private List<string> GetHeaders(ExcelWorksheet sheet)
        {
            List<string> headers = new();

            for (int i = 1; i <= sheet.Dimension.Columns; i++)
            {
                headers.Add(sheet.Cells[1, i].Text);
            }
            return headers;
        }

        /// <summary>
        /// Retrieves single row of data from XLSX file
        /// </summary>
        /// <param name="sheet">A <see cref="ExcelWorksheet"/> from which data will be read</param>
        /// <param name="row">An <see cref="int"/> indicating row number</param>
        /// <returns>A list of <see cref="string"/> containing data from specified row</returns>
        private List<string> GetDataRow(ExcelWorksheet sheet, int row)
        {
            List<string> data = new();

            for (int col = 1; col <= sheet.Dimension.Columns; col++)
            {
                data.Add(sheet.Cells[row, col].Text);
            }

            return data;
        }

        /// <summary>
        /// Writes data to file in xlsx format
        /// </summary>
        /// <param name="filePath">A <see cref="string"/> representing path to the file</param>
        /// <param name="records">A list of <see cref="Record"/> to be written</param>
        /// <returns>A <see cref="ReadingResult{T}"/> indicating whether the operation was successful or not.
        /// Contains error message in case of any error</returns>
        public ReadingResult<bool> WriteToFile(string filePath, List<Record> records)
        {
            try
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

                package.SaveAs(filePath);

                return ReadingResult<bool>.Success(true);
            }
            catch (Exception ex)
            {
                return ReadingResult<bool>.Fail($"Error while writing to file: {ex.Message}");
            }
        }

        /// <summary>
        /// Writes headers to the excel file
        /// </summary>
        /// <param name="sheet">A <see cref="ExcelWorksheet"/> representing file to which headers will be written</param>
        /// <param name="headers">A list of <see cref="string"/> containg column headers</param>
        private void WriteHeaders(ExcelWorksheet sheet, List<string> headers)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                sheet.Cells[1, i + 1].Value = headers[i];
            }
        }

        /// <summary>
        /// Writes single row of data to the excel file
        /// </summary>
        /// <param name="sheet">A <see cref="ExcelWorksheet"/> representing file to which headers will be written</param>
        /// <param name="data">A list of <see cref="string"/> data to be written on the row</param>
        /// <param name="row">An <see cref="int"/> indicating row number</param>
        private void WriteData(ExcelWorksheet sheet, List<string> data, int row)
        {
            for (int i = 0; i < data.Count; i++)
            {
                sheet.Cells[row, i + 1].Value = data[i];
            }
        }
    }
}