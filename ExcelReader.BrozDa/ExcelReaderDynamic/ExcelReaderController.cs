using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderDynamic
{
    internal class ExcelReaderController
    {
        public ExcelReader reader;
        public ExcelReaderRepository repo;
        public ExcelReaderController(ExcelReader reader, ExcelReaderRepository repo)
        {
            this.reader = reader;
            this.repo = repo;
        }

        public async void Run()
        {
            var records = reader.GetAllRecords();

            var headers = records[0].Headers;

            await repo.CreateTable(headers);

            await repo.InsertBulk(records);
        }
    }
}
