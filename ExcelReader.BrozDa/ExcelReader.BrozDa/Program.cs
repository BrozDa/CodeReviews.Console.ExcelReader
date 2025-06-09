using ExcelReader.BrozDa.Data;
using ExcelReader.BrozDa.Models;
using ExcelReader.BrozDa.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
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

            var services = new ServiceCollection();


            var sp = BuildServices(services);

            var db = sp.GetRequiredService<ExcelReaderContext>();
            db.Database.EnsureDeleted();
            db.Database.EnsureCreated();

            var controller = sp.GetRequiredService<ExcelReaderController>();
            controller.Run();

            List<string> dataProperties = typeof(Person).GetProperties().Select(x => x.Name).ToList();

            Console.WriteLine(string.Join("|", dataProperties));


        }
        public static ServiceProvider BuildServices(ServiceCollection services)
        {
            string path = Path.Combine(Environment.CurrentDirectory, "Resources/people-100.xlsx");

            services.AddDbContext<ExcelReaderContext>();
            services.AddScoped<ExcelReadingService>(sp => new ExcelReadingService(path));
            services.AddScoped<UiService>();
            services.AddScoped<ExcelReaderController>(sp => new ExcelReaderController(
                sp.GetRequiredService<ExcelReaderContext>(),
                sp.GetRequiredService<ExcelReadingService>(),
                sp.GetRequiredService<UiService>()
                ));

            return services.BuildServiceProvider();            
        }
    }
}
