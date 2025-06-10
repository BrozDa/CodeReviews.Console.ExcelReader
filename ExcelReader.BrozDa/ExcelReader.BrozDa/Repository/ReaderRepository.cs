using Dapper;
using ExcelReader.Brozda.Models;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

namespace ExcelReaderDynamic.Repository
{
    /// <summary>
    /// Manages access to the database
    /// </summary>
    internal class ReaderRepository
    {

        private string _connectionStringMasterDb;

        private string _connectionStringExcelReaderDB;

        /// <summary>
        /// Initializes a new instance of <see cref="ReaderRepository"/>
        /// </summary>
        /// <param name="config">An instance <see cref="IConfiguration"/></param>
        /// <remarks>An <see cref="IConfiguration"/> contains JSON appconfig containing two connection strings
        /// One for master db ("MasterDB") used to check if the DB exists, second for the DB where will be data stored ("ExcelReaderDB")</remarks>
        public ReaderRepository(IConfiguration config)
        {
            _connectionStringMasterDb = config.GetConnectionString("MasterDB") ?? string.Empty;
            _connectionStringExcelReaderDB = config.GetConnectionString("ExcelReaderDB") ?? string.Empty;
        }
        /// <summary>
        /// Initializes database - make sure the old DB is deleted (if it existed) ad created new one
        /// </summary>
        /// <returns> A Task result contains <see cref="ReadingResult{T}"/> indicating that both, delete and create operations were successful</returns>
        public async Task<RepositoryResult<bool>> InitializeDb()
        {
            var deleteResult = await EnsureDeleted();
            var createResult = await EnsureCreated();

            return new RepositoryResult<bool>() 
            { 
                IsSuccessful = deleteResult.IsSuccessful && createResult.IsSuccessful 
            };
        }
        /// <summary>
        /// Ensures that old database for data read from file is deleted.
        /// </summary>
        /// <returns>A Task result contains <see cref="ReadingResult{T}"/> indicating whether the database was deleted successfully </returns>
        public async Task<RepositoryResult<bool>> EnsureDeleted() 
        {
            var sql = @"IF EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        ALTER DATABASE [ExcelReaderDynamic] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                        DROP DATABASE [ExcelReaderDynamic];
                       END";

            return await ExecuteAsync(sql, _connectionStringMasterDb);
        }
        /// <summary>
        /// Ensures that new database for data read from file is created.
        /// </summary>
        /// <returns>A Task result contains <see cref="ReadingResult{T}"/> indicating whether the database was created successfully </returns>
        public async Task<RepositoryResult<bool>> EnsureCreated() 
        {
            var sql = @"IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'ExcelReaderDynamic')
                       BEGIN
                        CREATE DATABASE ExcelReaderDynamic
                       END";

            return await ExecuteAsync(sql, _connectionStringMasterDb);
        }

        /// <summary>
        /// Ensures that new table for data read from file is created successfully
        /// </summary>
        /// <param name="headers">A list of <see cref="string"/> containing column names for new table</param>
        /// <returns>A Task result contains <see cref="ReadingResult{T}"/> indicating whether the table was created successfully </returns>
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

        /// <summary>
        /// Inserts list of records to the database
        /// </summary>
        /// <param name="records">A list of <see cref="Record"/> to be inserted</param>
        /// <returns>A Task result contains <see cref="ReadingResult{T}"/> indicating whether records were successfully inserted</returns>        
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
        /// <summary>
        /// Retrieves all records stored in the database
        /// </summary>
        /// <returns>A Task result contains List of <see cref="Records"/> retrieved from the database, or a error message indicating faced issue</returns>
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
        /// <summary>
        /// Maps dynamic records retrieved from database to <see cref="Record"/> objects
        /// </summary>
        /// <param name="dataFromDb">An <see cref="IEnumerable{T}"/> of dynamic objects</param>
        /// <returns>A list of <see cref="Record"/> objects mapped from dynamic objects retrieved from the database"/></returns>
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
        /// <summary>
        /// Executes a SQL command passed againt database based on connection string
        /// </summary>
        /// <param name="sql">A string representation of SQL command</param>
        /// <param name="connectionstring">A connection string for the database</param>
        /// <returns>A Task result contains <see cref="RepositoryResult{T}"/> indicating if the execution was successfull</returns>
        private async Task<RepositoryResult<bool>> ExecuteAsync(string sql, string connectionstring)
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
