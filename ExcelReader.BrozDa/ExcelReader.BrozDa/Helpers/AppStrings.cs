namespace ExcelReader.Brozda.Helpers
{
    /// <summary>
    /// Represents strings used by the application
    /// </summary>
    internal static class AppStrings
    {
        public const string Menu_SpecifyFile = "Specify File";
        public const string Menu_ImportDataFromFile = "Import data from file";
        public const string Menu_ReadDatabase = "Read data from repository";
        public const string Menu_ExportDataToFile = "Export data to file";
        public const string Menu_Exit = "Exit the app";

        public const string Status_InitializingDb = "Initializing database";
        public const string Status_ReadingFile = "Reading File";
        public const string Status_CreatingTable = "Creating table";
        public const string Status_PopulatingDb = "Inserting records to database";
        public const string Status_DbPopulated = "Data successfully imported to database";
        public const string Status_OperationCancelled = "Operation Cancelled";
        public const string Status_GettingRecordsFromDb = "Reading database";
        public const string Status_WritingToFile = "Writing to file";
        public const string Status_WritingToFileFinished = "Writing to file finished";

        public const string Error_PathNotSpecified = "Path not specified";
        public const string Error_FileDoesNotExist = "Path invalid, File does not exist";

        public const string Io_PressAnyKeyToContinue = "Press any key to continue...";
        public const string Io_UnspecifiedError = "Unknown Error";
    }
}