using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Identity.Client;
namespace ExcelReaderDynamic
{
    internal class ExcelReaderRepository
    {

        string connectionStringDBcheck = @"Data Source=(localdb)\LOCALDB;Initial Catalog=master;Integrated Security=True;";

        string connectionStringDB = @"Data Source=(localdb)\LOCALDB;Initial Catalog=ExcelReaderDynamic;Integrated Security=True;";

        public async Task SetupDb()
        {
            await EnsureDeleted();
            await EnsureCreated();  
        }
        public async Task EnsureDeleted() 
        {
            var sql = @"IF EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        DROP DATABASE ExcelReaderDynamic
                       END";

            await ExecuteAsync(sql, connectionStringDBcheck);
        }
        public async Task EnsureCreated() 
        {
            var sql = @"IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        CREATE DATABASE ExcelReaderDynamic
                       END";

            await ExecuteAsync(sql, connectionStringDBcheck);
        }

        public async Task ExecuteAsync(string sql, string connectionstring)
        {
            using var connection = new SqlConnection(connectionstring);

            await connection.ExecuteAsync(sql);
        }
        public async Task CreateTable(List<string> headers)
        {
            string tableDefinition = "";
            foreach (var header in headers)
            {
                tableDefinition += $"[{header}] NVARCHAR(50),";
            }
            tableDefinition = tableDefinition.TrimEnd(',');
            string sql = @$"CREATE TABLE [ExcelData] ({tableDefinition});";

            Console.WriteLine(@$"CREATE TABLE [ExcelData] ({tableDefinition});");
            await ExecuteAsync(sql, connectionStringDB);
            Console.WriteLine("Done");
        }
        public async Task InsertBulk(List<Record> ExcelRecords)
        {
            using var connection = new SqlConnection(connectionStringDB);


            var columns = string.Join(",", ExcelRecords[0].Headers.Select(x => $"[{x}]"));

            foreach (var record in ExcelRecords) 
            {
                var data = string.Join(",", record.Data.Select(x => $"'{x}'"));

                string sql = @$"INSERT INTO [ExcelData] ({columns}) VALUES ({data});";

                await connection.ExecuteAsync(sql);
            }

        }
        public async Task TaskInsertRow(Record record)
        {
           
        }

    }

    

}
