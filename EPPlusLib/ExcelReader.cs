﻿using System.Collections.Generic;
using System.IO;
using Domain.Models;
using Domain.Ports;
using OfficeOpenXml;

namespace EPPlusLib
{
    public class ExcelReader : IExcelReader
    {
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

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (!IsValid(worksheet))
                {
                    throw new InvalidDataException($"Invalid file {file}");
                }

                var products = GetProducts(worksheet);
                return new Stock(products);
            }
        }

        private static bool IsValid(ExcelWorksheet worksheet)
        {
            var rows = GetRows(worksheet);
            var columns = GetColumns(worksheet);
            return rows > 1 
                   && columns == 3 
                   && string.Equals(worksheet.Name, Constants.WorksheetName);
        }

        private static ICollection<Product> GetProducts(ExcelWorksheet worksheet)
        {
            var products = new List<Product>();

            for (var row = Constants.StartRow; row <= GetRows(worksheet); row++)
            {
                var name = worksheet.GetValue<string>(row, Constants.NameColumn);
                var price = worksheet.GetValue<decimal>(row, Constants.PriceColumn);
                var quantity = worksheet.GetValue<int>(row, Constants.QuantityColumn);
                var product = new Product
                {
                    Name = name,
                    Price = price,
                    Quantity = quantity
                };
                products.Add(product);
            }

            return products;
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
