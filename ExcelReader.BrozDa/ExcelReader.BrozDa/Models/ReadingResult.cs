using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Brozda.Models
{
    internal class ReadingResult<T>
    {
        public bool IsSuccessful;
        public T? Data;
        public string? ErrorMessage;

        public static ReadingResult<T> Success(T data)
        {
            return new ReadingResult<T>()
            {
                IsSuccessful = true,
                Data = data
            };
        }
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
