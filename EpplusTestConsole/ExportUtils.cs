using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace EpplusTestConsole
{
    public static class ExportUtils
    {
        public static void LoadFromCollection<TObject>(ExcelWorksheet sheet, (string literal, int numeric) address, TObject[] collection, string[] colums, bool printHeaders = false) where TObject : class
        {
            //Передать данные сразу из двух массивов?
            MemberInfo[] membersToInclude = typeof(TObject)
           .GetProperties(BindingFlags.Instance | BindingFlags.Public)
           .Where(p => colums.Contains(p.Name))
           .ToArray();


            //Загрузка данных коллекции по адресу
            using var range = sheet.Cells[address.literal + address.numeric].LoadFromCollection(collection, printHeaders,
                OfficeOpenXml.Table.TableStyles.None,
                BindingFlags.Instance | BindingFlags.Public,
                membersToInclude);
            if (collection.Length > 0)
            {

                range.ProcessCommonStyles();
                range.ProcessTypeStyles(collection, printHeaders);
                //range.ProcessAttributes(collection, printHeaders);
            }
        }

        public static void BaseExport<TObject>(TObject[] objects, (string literal, int numeric) address, string fileName, string worksheetName, string[] colums, string[] headers)
            where TObject : class
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var file = new FileInfo(fileName);
            // Удаляет файл, если он уже существует
            if (file.Exists)
            {
                file.Delete();
            }
            using var package = new ExcelPackage(file);
            using var worksheet = package.Workbook.Worksheets.Add(worksheetName);

            if (worksheet != null)
            {
                worksheet.Cells[address.literal + address.numeric].LoadFromArrays(new List<string[]>(new[] { headers }));
                address.numeric += 1;

                LoadFromCollection(worksheet, address, objects, colums);
                worksheet.Cells.AutoFitColumns();
                worksheet.View.ZoomScale = 85;
                package.SaveAs(new FileInfo(fileName));
            }
        }
    }
}