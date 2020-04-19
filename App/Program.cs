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
            services.AddSingleton<IExcelWriter, ExcelWriter>();
            services.Configure<Settings>(configuration.GetSection(nameof(Settings)));

            var serviceProvider = services.BuildServiceProvider();

            var excelReaderFile = serviceProvider.GetService<IOptions<Settings>>().Value.ExcelReaderFilePath;
            var excelReader = serviceProvider.GetService<IExcelReader>();
            var stock = excelReader.Read(excelReaderFile);
            Console.WriteLine(stock);

            var excelWriterFile = serviceProvider.GetService<IOptions<Settings>>().Value.ExcelWriterFilePath;
            var excelWriter = serviceProvider.GetService<IExcelWriter>();
            var isGenerated = excelWriter.Write(excelWriterFile, stock) ? "is generated" : "is not generated";
            Console.WriteLine($"File '{excelWriterFile}' {isGenerated}");

            Console.WriteLine("Press any key to exit !");
            Console.ReadKey();
        }
    }
}
