using ExcelReader.BrozDa.Services;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Runtime.Versioning;

namespace ExcelReader.BrozDa
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ExcelPackage.License.SetNonCommercialPersonal("BrozDa");

            ExcelReadingService svc = new ExcelReadingService();

            var result = svc.ReadExcelFile();
            Console.WriteLine("a");
        }
    }
}
