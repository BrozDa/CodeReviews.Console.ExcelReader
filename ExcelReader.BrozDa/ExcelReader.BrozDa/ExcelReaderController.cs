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

        public ExcelReaderController(
            ExcelReaderContext dbContext, 
            ExcelReadingService readingService)
        {
            DbContext = dbContext;
            ReadingService = readingService;
        }

        public void Run()
        {
            List<Person> persons = GetExcelData(); //checked and list of 100 records is present here
            PopulateDb(persons);
        }
        private List<Person> GetExcelData()
        {
            Console.WriteLine("Reading Excel Data");
            return ReadingService.GetAllRecords();
        }
        private void PopulateDb(List<Person> persons)
        {
            Console.WriteLine("Populating DB");
            DbContext.AddRange(persons);
            DbContext.SaveChanges();
        }

    }
}
