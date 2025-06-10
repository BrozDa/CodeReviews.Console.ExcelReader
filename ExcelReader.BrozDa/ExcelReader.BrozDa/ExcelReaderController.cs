using ExcelReader.Brozda.Helpers;
using ExcelReader.Brozda.Models;
using ExcelReaderDynamic.Repository;
using ExcelReaderDynamic.Services;

namespace ExcelReaderDynamic
{
    /// <summary>
    /// Handles Excel/ CSV reader application flow
    /// </summary>
    internal class ExcelReaderController
    {
        /// <summary>
        /// Represents menu options
        /// </summary>
        public enum MenuOptions
        {
            SpecifyFile = 1,
            ImportDataFromFile,
            ReadDatabase,
            ExportDataToFile,
            Exit = 100
        }

        public Dictionary<int, string> menuOptionMap = new();
        public string FilePath { get; set; } = string.Empty;

        public ExcelReaderService ExcelReaderSvc { get; }
        public CsvReaderService CsvReaderSvc { get; }
        public ReaderRepository ReaderRepository { get; }
        private UiService UiSvc { get; }

        /// <summary>
        /// Intializes new instance of <see cref="ExcelReaderController"/>
        /// </summary>
        /// <param name="excelReader">An instance of <see cref="ExcelReaderService"/></param>
        /// <param name="csvReader">An instance of <see cref="CsvReaderService"/></param>
        /// <param name="repository">An instance of <see cref="ExcelReaderDynamic.Repository.ReaderRepository"/></param>
        /// <param name="ui">An instance of <see cref="UiService"/></param>
        public ExcelReaderController(ExcelReaderService excelReader,
            CsvReaderService csvReader,
            ReaderRepository repository,
            UiService ui
            )
        {
            ExcelReaderSvc = excelReader;
            CsvReaderSvc = csvReader;
            ReaderRepository = repository;
            UiSvc = ui;
        }

        /// <summary>
        /// Maps int representation of <see cref="MenuOptions"/> to its string representation
        /// </summary>
        private void MapMenu()
        {
            menuOptionMap.Add((int)MenuOptions.SpecifyFile, AppStrings.Menu_SpecifyFile);
            menuOptionMap.Add((int)MenuOptions.ImportDataFromFile, AppStrings.Menu_ImportDataFromFile);
            menuOptionMap.Add((int)MenuOptions.ReadDatabase, AppStrings.Menu_ReadDatabase);
            menuOptionMap.Add((int)MenuOptions.ExportDataToFile, AppStrings.Menu_ExportDataToFile);
            menuOptionMap.Add((int)MenuOptions.Exit, AppStrings.Menu_Exit);
        }

        /// <summary>
        /// Maps menu options, Initializes database and starts application flow
        /// </summary>
        /// <returns></returns>
        public async Task Run()
        {
            MapMenu();

            var isDbInitialized = await InitializeDb();

            if (!isDbInitialized)
            {
                Environment.Exit(-1);
            }
            UiSvc.ClearConsole();

            var input = UiSvc.GetMenuInput(menuOptionMap);

            while (input != (int)MenuOptions.Exit)
            {
                await ProcessUserChoice(input);
                Console.Clear();
                input = UiSvc.GetMenuInput(menuOptionMap);
            }
        }

        /// <summary>
        /// Intializes the database, deleting the old one if it exists and creating a fresh and empty one
        /// </summary>
        /// <returns>A Task result contains <see cref="bool"/> indicating success or fail</returns>
        private async Task<bool> InitializeDb()
        {
            UiSvc.PrintText(AppStrings.Status_InitializingDb);

            var initializeResult = await ReaderRepository.InitializeDb();

            if (!initializeResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg($"Db not initialized! Please contact admin.\n{initializeResult.ErrorMessage}");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Gets user choice and calls corespon
        /// </summary>
        /// <param name="choice">An int representation of <see cref="MenuOptions"/></param>
        private async Task ProcessUserChoice(int choice)
        {
            switch (choice)
            {
                case (int)MenuOptions.SpecifyFile:
                    ProcessSpecifyFile();
                    break;

                case (int)MenuOptions.ImportDataFromFile:
                    await ProcessImportDataFromFile();
                    break;

                case (int)MenuOptions.ReadDatabase:
                    await ProcessReadDatabase();
                    break;

                case (int)MenuOptions.ExportDataToFile:
                    await ProcessExportDataToFile();
                    break;
            }
        }

        /// <summary>
        /// Gets valid file from the user and stores it to FilePath property
        /// </summary>
        private void ProcessSpecifyFile()
        {
            string path = UiSvc.GetFilePathFromUser();

            FilePath = path;

            if (!File.Exists(FilePath))
            {
                UiSvc.PrintText(AppStrings.Error_FileDoesNotExist);
                UiSvc.PressAnyKeyToContinue();
            }
        }

        /// <summary>
        /// Reads data from file and if the reading was successful insert retrieved data to the database
        /// </summary>
        private async Task ProcessImportDataFromFile()
        {
            var records = GetDataFromFile();
            if (records == null) { return; }

            await PopulateDatabase(records);

            UiSvc.PressAnyKeyToContinue();
        }

        /// <summary>
        /// Creates table and inserts data to the new table
        /// </summary>
        /// <param name="records">A list of <see cref="Record"/> to be inserted</param>
        private async Task PopulateDatabase(List<Record> records)
        {
            var headers = records[0].Headers;

            UiSvc.PrintText(AppStrings.Status_CreatingTable);
            var createTableResult = await ReaderRepository.CreateTable(headers);

            if (!createTableResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg(createTableResult.ErrorMessage);
                return;
            }

            UiSvc.PrintText(AppStrings.Status_PopulatingDb);

            var insertResult = await ReaderRepository.InsertBulk(records);

            if (!insertResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg(createTableResult.ErrorMessage);
                return;
            }

            UiSvc.PrintText(AppStrings.Status_DbPopulated);
        }

        /// <summary>
        /// Checks if path stored in FilePath property is valid and if so, retrieves data from the file
        /// </summary>
        /// <returns></returns>
        private List<Record>? GetDataFromFile()
        {
            if (FilePath == string.Empty)
            {
                UiSvc.PrintText(AppStrings.Error_PathNotSpecified);
                UiSvc.PressAnyKeyToContinue();
                return null;
            }
            if (!File.Exists(FilePath))
            {
                UiSvc.PrintText(AppStrings.Error_FileDoesNotExist);
                UiSvc.PressAnyKeyToContinue();
                return null;
            }

            UiSvc.PrintText(AppStrings.Status_ReadingFile);

            return GetRecords();
        }

        /// <summary>
        /// Checks file extensions and calls respective service to read data from the file
        /// </summary>
        /// <returns>A list of <see cref="Record"/> retrieved from file.
        /// Error message is printed in case of any error and null is returned</returns>
        private List<Record>? GetRecords()
        {
            string extension = Path.GetExtension(FilePath).ToLowerInvariant();

            ReadingResult<List<Record>> result;

            switch (extension)
            {
                case ".xlsx": result = ExcelReaderSvc.ReadAllRecords(FilePath); break;
                case ".csv": result = CsvReaderSvc.ReadAllRecords(FilePath); break;
                default: result = ReadingResult<List<Record>>.Fail("Unsuported File"); break;
            }

            if (result.IsSuccessful)
            {
                return result.Data;
            }
            else
            {
                UiSvc.PrintErrorMsg(result.ErrorMessage);
                return null;
            }
        }

        /// <summary>
        /// Retrieves data from the database. Data is printed to the output if data is sucessfully retrieved.
        /// Respective error message is shown otherwise
        /// </summary>
        private async Task ProcessReadDatabase()
        {
            var readDbResult = await ReaderRepository.GetAll();

            if (!readDbResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg(readDbResult.ErrorMessage!);
                return;
            }

            UiSvc.PrintRecords(readDbResult.Data!);
        }

        /// <summary>
        /// Gets path where file should be created, gets confirmation of overwritting and writes data from database to the file
        /// Respective error message is shown in case of any issue.
        /// </summary>
        private async Task ProcessExportDataToFile()

        {
            string path = UiSvc.GetFilePathFromUser();

            if (File.Exists(path) && !UiSvc.ConfirmOverwrite())
            {
                UiSvc.PrintText(AppStrings.Status_OperationCancelled);
                return;
            }

            UiSvc.PrintText(AppStrings.Status_GettingRecordsFromDb);

            var readDbResult = await ReaderRepository.GetAll();

            if (!readDbResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg(readDbResult.ErrorMessage);
                UiSvc.PressAnyKeyToContinue();
                return;
            }

            string extension = Path.GetExtension(path).ToLowerInvariant();
            switch (extension)
            {
                case ".xlsx":
                    UiSvc.PrintText(AppStrings.Status_WritingToFile);
                    ExcelReaderSvc.WriteToFile(path, readDbResult.Data!);
                    break;

                case ".csv":
                    UiSvc.PrintText(AppStrings.Status_WritingToFile);
                    CsvReaderSvc.WriteToFile(path, readDbResult.Data!);
                    break;

                default: break;
            }
            UiSvc.PrintText(AppStrings.Status_WritingToFileFinished);
            UiSvc.PressAnyKeyToContinue();
        }
    }
}