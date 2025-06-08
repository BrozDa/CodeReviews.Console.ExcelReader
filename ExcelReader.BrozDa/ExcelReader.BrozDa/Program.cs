using ExcelReader.BrozDa.Services;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace ExcelReader.BrozDa
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ExcelPackage.License.SetNonCommercialPersonal("BrozDa");

            string path = Path.Combine(Environment.CurrentDirectory, "Resources/people-100.xlsx");

            ExcelReadingService service = new ExcelReadingService(path);

            var records = service.GetAllRecords();
            Console.ReadLine();


        }
    }
}
