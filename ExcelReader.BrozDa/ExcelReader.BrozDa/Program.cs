﻿using ExcelReaderDynamic.Repository;
using ExcelReaderDynamic.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;

namespace ExcelReaderDynamic
{
    internal class Program
    {
        private static async Task Main()
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
                .AddJsonFile(Path.Combine(Environment.CurrentDirectory, "appconfig.json"))
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