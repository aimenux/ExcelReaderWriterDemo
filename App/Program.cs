using System;
using System.IO;
using Domain.Ports;
using EPPlusLib;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace App
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var environment = Environment.GetEnvironmentVariable("ENVIRONMENT") ?? "DEV";

            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .AddJsonFile($"appsettings.{environment}.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();

            var services = new ServiceCollection();
            services.AddSingleton<IExcelReader, ExcelReader>();
            services.Configure<Settings>(configuration.GetSection(nameof(Settings)));

            var serviceProvider = services.BuildServiceProvider();
            var excelFile = serviceProvider.GetService<IOptions<Settings>>().Value.ExcelFilePath;
            var excelReader = serviceProvider.GetService<IExcelReader>();
            var stock = excelReader.Read(excelFile);
            Console.WriteLine(stock);

            Console.WriteLine("Press any key to exit !");
            Console.ReadKey();
        }
    }
}
