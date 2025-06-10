namespace ExcelReader.Brozda.Models
{
    /// <summary>
    /// Represent result of action performed by Reading service
    /// </summary>
    /// <typeparam name="T">Datatype contained in data portion of the result</typeparam>
    internal class ReadingResult<T>
    {
        public bool IsSuccessful;
        public T? Data;
        public string? ErrorMessage;

        /// <summary>
        /// Gets successful <see cref="ReadingResult{T}"/>
        /// </summary>
        /// <param name="data">Data to be included in the result</param>
        /// <returns>A successful <see cref="ReadingResult{T}"/></returns>
        public static ReadingResult<T> Success(T data)
        {
            return new ReadingResult<T>()
            {
                IsSuccessful = true,
                Data = data
            };
        }

        /// <summary>
        /// Gets failed <see cref="ReadingResult{T}"/>
        /// </summary>
        /// <param name="errorMessage">Error message indicating why the action failed</param>
        /// <returns>A <see cref="ReadingResult{T}"/> indicating failed operation</returns>
        public static ReadingResult<T> Fail(string errorMessage)
        {
            return new ReadingResult<T>()
            {
                IsSuccessful = true,
                ErrorMessage = errorMessage
            };
        }
    }
}