using ExcelReaderDynamic.Repository;
using ExcelReaderDynamic.Services;

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

        public string file = Environment.CurrentDirectory;
        public string FilePath { get; set; }

        public ExcelReaderService ExcelReaderSvc { get; set; }
        public CsvReaderService CsvReaderSvc { get; set; }
        public ReaderRepository ReaderRepository { get; set; }
        private UiService UiSvc { get; set; }

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
            await ReaderRepository.InitializeDb();

            var input = UiSvc.GetMenuInput(menuOptionMap);

            while (input != (int)MenuOptions.Exit) 
            {
                await ProcessUserChoice(input);
                Console.Clear();
                input = UiSvc.GetMenuInput(menuOptionMap);
                
            }
        }

        private async Task ProcessUserChoice(int choice)
        {
            switch (choice)
            { 
                case (int)MenuOptions.SpecifyFile: ProcessSpecifyFile(); break;
                case (int)MenuOptions.ImportDataFromFile: await ProcessImportDataFromFile(); break;
                case (int)MenuOptions.ReadDatabase: await ProcessReadDatabase(); break;
                case (int)MenuOptions.ExportDataToFile: await ProcessExportDataToFile();  break;
            }
        }
        private void ProcessSpecifyFile()
        {
            string path = UiSvc.GetFilePathFromUser();
            
            while (!File.Exists(path))
            {
                Console.WriteLine("File does not exist");
                path = UiSvc.GetFilePathFromUser();
            }

            FilePath = path;
        }
        private async Task ProcessImportDataFromFile()
        {
            if(FilePath == string.Empty || !File.Exists(FilePath))
            {
                Console.WriteLine("File not specified or the path is invalid");
            }

            Console.WriteLine("Reading file");
            var records = GetRecords();

            if(records == null) 
            {
                // handle error
                return;
            }

            var headers = records[0].Headers;

            Console.WriteLine("Creating table");
            await ReaderRepository.CreateTable(headers);

            Console.WriteLine("Inserting records");
            await ReaderRepository.InsertBulk(records);

            Console.WriteLine("Data successfully imported");

            UiSvc.PressAnyKeyToContinue();
        }

        private List<Record>? GetRecords()
        {
            string extension = Path.GetExtension(FilePath).ToLowerInvariant();

            return extension switch
            {
                ".xlsx" => ExcelReaderSvc.ReadAllRecords(FilePath),
                ".csv" => CsvReaderSvc.ReadAllRecords(FilePath),
                _ => null
            };
        }
        private async Task ProcessReadDatabase()
        {
            var recordsFromDB = await ReaderRepository.GetAll();
            UiSvc.PrintRecords(recordsFromDB);
            
        }
        private async Task ProcessExportDataToFile()
        
        {
            string path = UiSvc.GetFilePathFromUser();
            
            //overwrite logic
            var recordsFromDB = await ReaderRepository.GetAll();

            string extension = Path.GetExtension(path).ToLowerInvariant();
            switch (extension)
            {
                case ".xlsx": ExcelReaderSvc.WriteToFile(path, recordsFromDB); break;
                case ".csv": CsvReaderSvc.WriteToFile(path, recordsFromDB); break;
                default: break;
            }
            
            Console.WriteLine("e");
        }
    }
}
