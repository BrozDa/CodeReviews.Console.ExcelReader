using ExcelReader.BrozDa.Data;
using ExcelReader.BrozDa.Models;
using ExcelReader.BrozDa.Services;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;

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
