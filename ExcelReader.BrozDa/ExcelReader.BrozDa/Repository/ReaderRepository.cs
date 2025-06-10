using Dapper;
using ExcelReader.Brozda.Models;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Reflection.Metadata.Ecma335;
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
        public async Task<RepositoryResult<bool>> InitializeDb()
        {
            var deleteResult = await EnsureDeleted();
            var createResult = await EnsureCreated();

            return new RepositoryResult<bool>() 
            { 
                IsSuccessful = deleteResult.IsSuccessful && createResult.IsSuccessful 
            };
        }
        public async Task<RepositoryResult<bool>> EnsureDeleted() 
        {
            var sql = @"IF EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        ALTER DATABASE [ExcelReaderDynamic] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                        DROP DATABASE [ExcelReaderDynamic];
                       END";

            return await ExecuteAsync(sql, _connectionStringMasterDb);
        }
        public async Task<RepositoryResult<bool>> EnsureCreated() 
        {
            var sql = @"IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        CREATE DATABASE ExcelReaderDynamic
                       END";

            return await ExecuteAsync(sql, _connectionStringMasterDb);
        }


        public async Task<RepositoryResult<bool>> CreateTable(List<string> headers)
        {

            string tableDefinition = "";
            foreach (var header in headers)
            {
                tableDefinition += $"[{header}] NVARCHAR(50),";
            }
            tableDefinition = tableDefinition.TrimEnd(',');
            string sql = @$"CREATE TABLE [ExcelData] ({tableDefinition});";

            return await ExecuteAsync(sql, _connectionStringExcelReaderDB);

        }
            
            
        public async Task<RepositoryResult<bool>> InsertBulk(List<Record> records)
        {
            try
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

                    int affectedRows = await connection.ExecuteAsync(sql, sqlParameters);
                }

                return RepositoryResult<bool>.NonQuerrrySuccess();

            }
            catch (Exception ex) 
            {
                return RepositoryResult<bool>.NonQuerrryFail($"Error while accessing database: {ex.Message}");
            }
            
        }

        public async Task<RepositoryResult<List<Record>>> GetAll()
        {
            try
            {
                string sql = $"SELECT * FROM [ExcelData];";
                using var connection = new SqlConnection(_connectionStringExcelReaderDB);

                var rows = await connection.QueryAsync(sql);

                
                var recordList = MapDynamicToRecord(rows);

                return new RepositoryResult<List<Record>>()
                {
                    IsSuccessful = true,
                    Data = recordList
                };
            }
            catch (Exception ex) 
            {
                return new RepositoryResult<List<Record>>()
                {
                    IsSuccessful = false,
                    ErrorMessage = $"Error while accessing database: {ex.Message}"
                };
            }
            
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
        public async Task<RepositoryResult<bool>> ExecuteAsync(string sql, string connectionstring)
        {
            try
            {
                using var connection = new SqlConnection(connectionstring);
                await connection.ExecuteAsync(sql);

                return RepositoryResult<bool>.NonQuerrrySuccess();
            }
            catch (Exception ex) 
            {
                return RepositoryResult<bool>.NonQuerrryFail($"Error while accessing database: {ex.Message}");
            }
                
 
        }



    }

    

}
