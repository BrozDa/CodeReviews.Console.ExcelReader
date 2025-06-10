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
            menuOptionMap.Add((int)MenuOptions.SpecifyFile, "Specify File");
            menuOptionMap.Add((int)MenuOptions.ImportDataFromFile, "Import data from file");
            menuOptionMap.Add((int)MenuOptions.ReadDatabase, "Read data from repository");
            menuOptionMap.Add((int)MenuOptions.ExportDataToFile, "Export data to file");
            menuOptionMap.Add((int)MenuOptions.Exit, "Exit the app");
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
            Console.WriteLine("Initializing DB");

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

            Console.WriteLine("Creating table");
            var createTableResult = await ReaderRepository.CreateTable(headers);
            if (!createTableResult.IsSuccessful)
            {

                UiSvc.PrintErrorMsg(createTableResult.ErrorMessage);
                return;
            }

            Console.WriteLine("Inserting records");
            var insertResult = await ReaderRepository.InsertBulk(records);

            if (!insertResult.IsSuccessful)
            {
                UiSvc.PrintErrorMsg(createTableResult.ErrorMessage);
                return;
            }

            Console.WriteLine("Data successfully imported");
        }

        private List<Record>? GetDataFromFile()
        {
            if (FilePath == string.Empty)
            {
                Console.WriteLine("Path not specified");
                UiSvc.PressAnyKeyToContinue();
                return null;
            }
            if (!File.Exists(FilePath))
            {
                Console.WriteLine("Path invalid, File does not exist");
                UiSvc.PressAnyKeyToContinue();
                return null;
            }


            Console.WriteLine("Reading file");

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
               Console.WriteLine("Operation cancelled");
                return;
            }

            var readDbResult = await ReaderRepository.GetAll();

            if (!readDbResult.IsSuccessful)
            {
                Console.WriteLine(readDbResult.ErrorMessage);
                return;
            }

            string extension = Path.GetExtension(path).ToLowerInvariant();
            switch (extension)
            {
                case ".xlsx": ExcelReaderSvc.WriteToFile(path, readDbResult.Data!); break;
                case ".csv": CsvReaderSvc.WriteToFile(path, readDbResult.Data!); break;
                default: break;
            }
           
        }
    }
}
