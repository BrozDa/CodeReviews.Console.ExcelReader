using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Reflection.PortableExecutable;
using static System.Runtime.InteropServices.JavaScript.JSType;
namespace ExcelReaderDynamic.Repository
{
    internal class ReaderRepository
    {

        private string _connectionStringMasterDb;

        private string _connectionStringExcelReaderDB;
        public ReaderRepository(IConfiguration config)
        {
            _connectionStringMasterDb = config.GetConnectionString("MasterDB") ?? string.Empty;
            _connectionStringExcelReaderDB = config.GetConnectionString("ExcelReaderDB") ?? string.Empty;
        }
        public async Task InitializeDb()
        {
            await EnsureDeleted();
            await EnsureCreated();  
        }
        public async Task EnsureDeleted() 
        {
            var sql = @"IF EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        ALTER DATABASE [ExcelReaderDynamic] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                        DROP DATABASE [ExcelReaderDynamic];
                       END";

            await ExecuteAsync(sql, _connectionStringMasterDb);
        }
        public async Task EnsureCreated() 
        {
            var sql = @"IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        CREATE DATABASE ExcelReaderDynamic
                       END";

            await ExecuteAsync(sql, _connectionStringMasterDb);
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

            await ExecuteAsync(sql, _connectionStringExcelReaderDB);
        }
        public async Task InsertBulk(List<Record> records)
        {
            using var connection = new SqlConnection(_connectionStringExcelReaderDB);

            var headers = records[0].Headers;
            var normalizedHeaders = headers.Select(x => $"{x.Replace(" ", "")}").ToList();

            var columnsString = string.Join(",", headers.Select(x => $"[{x}]"));
            var parametersString = string.Join(",", normalizedHeaders.Select(x => $"@{x}"));

            string sql = @$"INSERT INTO [ExcelData] ({columnsString}) VALUES ({parametersString});";

            foreach (var record in records)
            {
                var sqlParameters = new DynamicParameters();
                for (int i = 0; i < headers.Count; i++)
                {
                    sqlParameters.Add($"@{normalizedHeaders[i]}", record.Data[i]);
                }

                await connection.ExecuteAsync(sql, sqlParameters);
            }
        }

        public async Task<List<Record>> GetAll()
        {
            string sql = $"SELECT * FROM [ExcelData];";
            using var connection = new SqlConnection(_connectionStringExcelReaderDB);

            var rows = await connection.QueryAsync(sql);


            return MapDynamicToRecord(rows);
        }
        
        private List<Record> MapDynamicToRecord(IEnumerable<dynamic> dataFromDb)
        {
            List<Record> records = new();
            foreach (dynamic dataRow in dataFromDb) 
            {
                var dict = (IDictionary<string, object>)dataRow;

                var headers = dict.Keys.ToList();
                var data = dict.Values.Select(v => v.ToString() ?? string.Empty).ToList();

                records.Add(new Record {Headers = headers, Data = data ?? new List<string>()});
            }
            return records;
        }
        public async Task ExecuteAsync(string sql, string connectionstring)
        {

            using var connection = new SqlConnection(connectionstring);

            await connection.ExecuteAsync(sql);
        }



    }

    

}
