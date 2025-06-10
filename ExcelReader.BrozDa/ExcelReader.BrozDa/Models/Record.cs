namespace ExcelReaderDynamic
{
    /// <summary>
    /// Represents data row from file
    /// </summary>
    internal class Record
    {
        public List<string> Headers { get; set; } = new();
        public List<string> Data { get; set; } = new();
    }
}