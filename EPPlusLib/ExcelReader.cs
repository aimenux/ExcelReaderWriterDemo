using System.Collections.Generic;
using System.IO;
using Domain.Models;
using Domain.Ports;
using OfficeOpenXml;

namespace EPPlusLib
{
    public class ExcelReader : IExcelReader
    {
        private const int StartRow = 2;
        private const int NameColumn = 1;
        private const int PriceColumn = 2;
        private const int QuantityColumn = 3;
        private const string WorksheetName = "Stock";

        static ExcelReader()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
        }

        public Stock Read(string file)
        {
            var fileInfo = new FileInfo(file);
            if (!fileInfo.Exists)
            {
                throw new FileNotFoundException($"Unfound file '{file}'");
            }

            var products = new List<Product>();

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (!IsValid(worksheet))
                {
                    throw new InvalidDataException($"Invalid file {file}");
                }

                var rows = GetRows(worksheet);
                for (var row = StartRow; row <= rows; row++)
                {
                    var name = worksheet.GetValue<string>(row, NameColumn);
                    var price = worksheet.GetValue<decimal>(row, PriceColumn);
                    var quantity = worksheet.GetValue<int>(row, QuantityColumn);
                    var product = new Product
                    {
                        Name = name,
                        Price = price,
                        Quantity = quantity
                    };
                    products.Add(product);
                }
            }
            
            return new Stock(products);
        }

        private static bool IsValid(ExcelWorksheet worksheet)
        {
            var rows = GetRows(worksheet);
            var columns = GetColumns(worksheet);
            return rows > 1 
                   && columns == 3 
                   && string.Equals(worksheet.Name, WorksheetName);
        }

        private static int GetRows(ExcelWorksheet worksheet)
        {
            return worksheet.Dimension.End.Row - worksheet.Dimension.Start.Row + 1;
        }

        private static int GetColumns(ExcelWorksheet worksheet)
        {
            return worksheet.Dimension.End.Column - worksheet.Dimension.Start.Column + 1;
        }
    }
}
