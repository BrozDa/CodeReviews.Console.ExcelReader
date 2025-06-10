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

        public bool ImportDbToCsv(List<Record> records)
        {
            throw new NotImplementedException();
        }
    }
}
