namespace ExcelReader.Brozda.Models
{
    internal class RepositoryResult<T>
    {
        public bool IsSuccessful;
        public T? Data;
        public string? ErrorMessage;

        public static RepositoryResult<bool> NonQuerrrySuccess()
        {
            return new RepositoryResult<bool> { IsSuccessful = true };

        }
        public static RepositoryResult<bool> NonQuerrryFail(string errorMessage)
        {
            return new RepositoryResult<bool> { IsSuccessful = true, ErrorMessage = errorMessage};
        }
    }

    
}
