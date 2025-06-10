using CsvHelper;
using ExcelReader.Brozda.Models;
using System.Globalization;

namespace ExcelReaderDynamic.Services
{
    /// <summary>
    /// Handles reading from and writing to CSV file
    /// </summary>
    internal class CsvReaderService
    {
        /// <summary>
        /// Reads all records from CSV file
        /// </summary>
        /// <param name="filePath">A <see cref="string"/> representing path to the file</param>
        /// <returns>A <see cref="ReadingResult{T}"/> indicating wheter the operation was successful
        /// A <see cref="ReadingResult{T}"/> contains retrieved data in form of List of <see cref="Record"/> in case of success
        /// or contains error message in case of any error</returns>
        public ReadingResult<List<Record>> ReadAllRecords(string filePath)
        {
            try
            {
                using StreamReader reader = new StreamReader(filePath);

                using var csvReader = new CsvReader(reader, CultureInfo.InvariantCulture);

                csvReader.Read();
                csvReader.ReadHeader();

                List<string> headers = GetHeaders(csvReader);
                List<Record> records = new();

                while (csvReader.Read())
                {
                    records.Add(new Record
                    {
                        Headers = headers,
                        Data = csvReader.Parser.Record?.ToList() ?? new List<string>()
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
        /// Retrieves column headers from CSV file
        /// </summary>
        /// <param name="reader">A <see cref="CsvReader"/> used to read the file</param>
        /// <returns>A list of <see cref="string"/> containing headers from file</returns>
        private List<string> GetHeaders(CsvReader reader) => reader.HeaderRecord?.ToList() ?? new List<string>();

        /// <summary>
        /// Writes data to file in csv format
        /// </summary>
        /// <param name="filePath">A <see cref="string"/> representing path to the file</param>
        /// <param name="records">A list of <see cref="Record"/> to be written</param>
        /// <returns>A <see cref="ReadingResult{T}"/> indicating whether the operation was successful or not.
        /// Contains error message in case of any error</returns>
        public ReadingResult<bool> WriteToFile(string filePath, List<Record> records)
        {
            try
            {
                using StreamWriter writer = new StreamWriter(filePath);
                using var csvWriter = new CsvWriter(writer, CultureInfo.InvariantCulture);

                foreach (var header in records[0].Headers)
                {
                    csvWriter.WriteField(header);
                }

                csvWriter.NextRecord();

                foreach (var record in records)
                {
                    foreach (var field in record.Data)
                    {
                        csvWriter.WriteField(field);
                    }
                    csvWriter.NextRecord();
                }

                return ReadingResult<bool>.Success(true);
            }
            catch (Exception ex)
            {
                return ReadingResult<bool>.Fail($"Error while writing to file: {ex.Message}");
            }
        }
    }
}