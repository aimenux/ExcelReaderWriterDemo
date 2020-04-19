using System;
using System.IO;
using Domain;
using Domain.Models;
using Domain.Ports;
using OfficeOpenXml;

namespace EPPlusLib
{
    public class ExcelWriter : IExcelWriter
    {
        static ExcelWriter()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
        }

        public bool Write(string file, Stock stock)
        {
            try
            {
                var fileInfo = new FileInfo(file);
                if (fileInfo.Exists)
                {
                    fileInfo.Delete();
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add(Constants.WorksheetName);
                    worksheet.Cells[nameof(Constants.A1)].Value = Constants.A1;
                    worksheet.Cells[nameof(Constants.B1)].Value = Constants.B1;
                    worksheet.Cells[nameof(Constants.C1)].Value = Constants.C1;
                    var row = Constants.StartRow;
                    foreach (var product in stock.Products)
                    {
                        worksheet.Cells[row, Constants.NameColumn].Value = product.Name;
                        worksheet.Cells[row, Constants.PriceColumn].Value = product.Price;
                        worksheet.Cells[row, Constants.QuantityColumn].Value = product.Quantity;
                        row++;
                    }
                    package.SaveAs(fileInfo);
                }

                return true;
            }
            catch (Exception ex)
            {
                ConsoleColor.Red.WriteLine(ex);
                return false;
            }
        }
    }
}
