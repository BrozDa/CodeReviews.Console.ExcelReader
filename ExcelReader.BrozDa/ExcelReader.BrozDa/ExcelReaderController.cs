using ExcelReader.Brozda.Helpers;
using ExcelReader.Brozda.Models;
using ExcelReaderDynamic.Repository;
using ExcelReaderDynamic.Services;
using Spectre.Console;
using System.IO;

namespace ExcelReaderDynamic
{
    internal class ExcelReaderController
    {
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
        public ReaderRepository ReaderRepository { get;}
        private UiService UiSvc { get; }

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
            MapMenu();
        }

        private void MapMenu()
        {
            menuOptionMap.Add((int)MenuOptions.SpecifyFile, AppStrings.Menu_SpecifyFile);
            menuOptionMap.Add((int)MenuOptions.ImportDataFromFile, AppStrings.Menu_ImportDataFromFile);
            menuOptionMap.Add((int)MenuOptions.ReadDatabase, AppStrings.Menu_ReadDatabase);
            menuOptionMap.Add((int)MenuOptions.ExportDataToFile, AppStrings.Menu_ExportDataToFile);
            menuOptionMap.Add((int)MenuOptions.Exit, AppStrings.Menu_Exit);
        }
        public async Task Run()
        {
            var isDbInitialized = await InitializeDb();

            if (!isDbInitialized ) 
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
        private async Task ProcessImportDataFromFile()
        {
            var records = GetDataFromFile();
            if (records == null) { return; }

            await PopulateDatabase(records);

            UiSvc.PressAnyKeyToContinue();
        }
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

            var records = GetRecords();


            return records;
        }
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
        private async Task ProcessExportDataToFile()
        
        {
            string path = UiSvc.GetFilePathFromUser();

            //overwrite logic
            if (File.Exists(path) && !UiSvc.ConfirmOverwrite())
            {
                UiSvc.PrintText(AppStrings.Status_OperationCancelled);
                return;
            }

            UiSvc.PrintText(AppStrings.Status_GettingRecordsFromDb);

            var readDbResult = await ReaderRepository.GetAll();

            if (!readDbResult.IsSuccessful)
            {
                UiSvc.PrintText(readDbResult.ErrorMessage);
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
