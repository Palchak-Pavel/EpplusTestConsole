using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace EpplusTestConsole
{
    public static class ExportUtils
    {
        public static void LoadFromCollection<TObject>(ExcelWorksheet sheet, (string literal, int numeric) address, TObject[] collection, bool printHeaders = false) where TObject : class
        {
            //Загрузка данных коллекции по адресу
            using var range = sheet.Cells[address.literal + address.numeric].LoadFromCollection(collection, printHeaders);
            if (collection.Length > 0)
            {
                range.ProcessCommonStyles();
                range.ProcessTypeStyles(collection, printHeaders);
                //range.ProcessAttributes(collection, printHeaders);
            }
        }
        
        public static void BaseExport<TObject>(TObject[] objects, (string literal, int numeric) address, string fileName, string worksheetName)
            where  TObject: class
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using var package = new ExcelPackage(new FileInfo(fileName));
            using var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            if (worksheet != null)
            {
                        
                LoadFromCollection(worksheet, address, objects, true);
                worksheet.Cells.AutoFitColumns();
                worksheet.View.ZoomScale = 85;
                package.SaveAs(new FileInfo(fileName));
            }
        }
    }
}