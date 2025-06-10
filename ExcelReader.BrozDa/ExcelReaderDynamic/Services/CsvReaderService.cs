using CsvHelper;
using System.Globalization;

namespace ExcelReaderDynamic.Services
{
    internal class CsvReaderService
    {
        public List<Record> ReadAllRecords(string filePath)
        {
            using StreamReader reader = new StreamReader(filePath);

            using var csvReader = new CsvReader(reader, CultureInfo.InvariantCulture);

            csvReader.Read();
            csvReader.ReadHeader();


            List<string> headers = GetHeaders(csvReader);
            List<Record> records = new ();

            while (csvReader.Read())
            {
                records.Add(new Record
                {
                    Headers = headers,
                    Data = csvReader.Parser.Record?.ToList() ?? new List<string>()
                });
            }

            return records;
        }
        private List<string> GetHeaders(CsvReader reader) => reader.HeaderRecord?.ToList() ?? new List<string>();

        public void WriteToFile(string path, List<Record> records)
        {
            using StreamWriter writer = new StreamWriter(path);
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

        }
    }
}
