namespace ExcelReader.Brozda.Models
{
    /// <summary>
    /// Represent result of action performed by Repository
    /// </summary>
    /// <typeparam name="T">Datatype contained in data portion of the result</typeparam>
    internal class RepositoryResult<T>
    {
        public bool IsSuccessful;
        public T? Data;
        public string? ErrorMessage;

        /// <summary>
        /// Gets a success <see cref="RepositoryResult{T}"/> for non-querry operation
        /// </summary>
        /// <returns>A <see cref="RepositoryResult{T}"/> indicating successful repository operation</returns>
        public static RepositoryResult<bool> NonQuerrrySuccess()
        {
            return new RepositoryResult<bool> { IsSuccessful = true };
        }

        /// <summary>
        /// Gets a failed <see cref="RepositoryResult{T}"/> for non-querry operation
        /// </summary>
        /// <param name="errorMessage">Error message indicating why the action failed</param>
        /// <returns>A <see cref="RepositoryResult{T}"/> indicating failed repository operation</returns>
        public static RepositoryResult<bool> NonQuerrryFail(string errorMessage)
        {
            return new RepositoryResult<bool> { IsSuccessful = true, ErrorMessage = errorMessage };
        }
    }
}