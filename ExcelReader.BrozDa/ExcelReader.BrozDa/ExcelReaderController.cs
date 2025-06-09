using ExcelReader.BrozDa.Data;
using ExcelReader.BrozDa.Models;
using ExcelReader.BrozDa.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.BrozDa
{
    internal class ExcelReaderController
    {
        public ExcelReaderContext DbContext { get; }
        public ExcelReadingService ReadingService { get; }

        public UiService UiService { get; }

        public ExcelReaderController(
            ExcelReaderContext dbContext, 
            ExcelReadingService readingService,
            UiService uiService)
        {
            DbContext = dbContext;
            ReadingService = readingService;
            UiService = uiService;
        }

       /* public void Run()
        {
            List<Person> persons = GetExcelData();
            PopulateDb(persons);

            UiService.PrintDatabase(GetPersonsFromDb());
        }*/
        /*private List<Person> GetExcelData()
        {
            Console.WriteLine("Reading Excel Data");
            return ReadingService.GetAllRecords();
        }
        private void PopulateDb(List<Person> persons)
        {
            Console.WriteLine("Populating DB");
            DbContext.Persons.AddRange(persons);
            DbContext.SaveChanges();
        }
        private List<Person> GetPersonsFromDb()
        {
            Console.WriteLine("Fetching data from db");
            return DbContext.Persons.ToList();
        }
        */
    }
}
