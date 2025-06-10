using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;
using ExcelReaderDynamic.Services;
using ExcelReaderDynamic.Repository;
using Microsoft.Extensions.Configuration;

namespace ExcelReaderDynamic
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            ExcelPackage.License.SetNonCommercialPersonal("BrozDa");

            var services = new ServiceCollection();

            var serviceProvider = BuildServices(services);

            var controller = serviceProvider.GetRequiredService<ExcelReaderController>();

            await controller.Run();

        }

        public static ServiceProvider BuildServices(ServiceCollection services)
        {
            var repoConfiguration = new ConfigurationBuilder()
                .AddJsonFile(@"E:\Git Repos\CodeReviews.Console.ExcelReader\ExcelReader.BrozDa\ExcelReaderDynamic\appconfig.json")
                .Build();

            services.AddSingleton<IConfiguration>(repoConfiguration);

            services.AddScoped<ExcelReaderService>();
            services.AddScoped<CsvReaderService>();
            services.AddScoped<ReaderRepository>();
            services.AddScoped<UiService>();
            services.AddScoped<ExcelReaderController>();
            

            return services.BuildServiceProvider();

        
        }
    }
}
