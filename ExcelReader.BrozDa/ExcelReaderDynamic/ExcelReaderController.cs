using ExcelReaderDynamic.Repository;
using ExcelReaderDynamic.Services;

namespace ExcelReaderDynamic
{
    internal class ExcelReaderController
    {
        public Dictionary<int, string> menuOptionMap = new();
        public enum MenuOptions
        {
            SpecifyFile = 1,
            ImportDataFromFile,
            ReadDatabase,
            ExportDataToFile,
            Exit = 100

        }

        public ExcelReaderService reader;
        public ExcelReaderRepository repo;
        private UiService _ui;
        public ExcelReaderController(ExcelReaderService reader, ExcelReaderRepository repo, UiService ui)
        {
            this.reader = reader;
            this.repo = repo;
            _ui = ui;
            MapMenu();
        }

        private void MapMenu()
        {
            menuOptionMap.Add((int)MenuOptions.SpecifyFile, "Specify File");
            menuOptionMap.Add((int)MenuOptions.ImportDataFromFile, "Import data from file");
            menuOptionMap.Add((int)MenuOptions.ReadDatabase, "Read data from repository");
            menuOptionMap.Add((int)MenuOptions.ExportDataToFile, "Export data from file");
            menuOptionMap.Add((int)MenuOptions.Exit, "Exit the app");
        }
        public async Task Run()
        {
            var input = _ui.GetMenuInput(menuOptionMap);

            while (input != (int)MenuOptions.Exit) 
            {
                await ProcessUserChoice(input);
                input = _ui.GetMenuInput(menuOptionMap);
            }
        }

        private async Task ProcessUserChoice(int choice)
        {
            switch (choice)
            { 
                case (int)MenuOptions.SpecifyFile: ProcessSpecifyFile(); break;
                case (int)MenuOptions.ImportDataFromFile: await ProcessImportDataFromFile(); break;
                case (int)MenuOptions.ReadDatabase: await ProcessReadDatabase(); break;
                case (int)MenuOptions.ExportDataToFile: break;
            }
        }
        private void ProcessSpecifyFile()
        {
            var path = _ui.GetFilePathFromUser();
            //string path = @"E:\Git Repos\CodeReviews.Console.ExcelReader\ExcelReader.BrozDa\ExcelReaderDynamic\people-100.xlsx";
            reader.FilePath = path;
        }
        private async Task ProcessImportDataFromFile()
        {
            if(reader.FilePath == string.Empty || !File.Exists(reader.FilePath))
            {
                Console.WriteLine("File not specified or the path is invalid");
            }
            Console.WriteLine("Reading Excel file");
            var records = reader.ReadFile();
            var headers = records[0].Headers;

            Console.WriteLine("Initialiazing DB");
            await repo.InitializeDb();

            Console.WriteLine("Creating table");
            await repo.CreateTable(headers);

            Console.WriteLine("Inserting records");
            await repo.InsertBulk(records);

            Console.WriteLine("Data successfully imported");
            Console.WriteLine("Press any key to continue");
            Console.ReadLine();
            Console.Clear();
        }

        private async Task ProcessReadDatabase()
        {
            var recordsFromDB = await repo.GetAll();
            _ui.PrintRecords(recordsFromDB);
        }
    }
}
