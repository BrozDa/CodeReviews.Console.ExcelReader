using OfficeOpenXml;

namespace ExcelReaderDynamic
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.License.SetNonCommercialPersonal("BrozDa");

            string path = @"E:\Git Repos\CodeReviews.Console.ExcelReader\ExcelReader.BrozDa\ExcelReaderDynamic\people-100.xlsx";

            ExcelReader reader = new ExcelReader(path);

            ExcelReaderRepository repo = new ExcelReaderRepository();
            await repo.SetupDb();

            ExcelReaderController ctrl = new ExcelReaderController(reader, repo);

            ctrl.Run();

            Console.ReadLine();
        }
    }
}
