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

        // Из-за разных наименований размеров у каждого товара не можем использовать стандартные названия полей
        // Поискать анонимный тип, в который можно запихнуть эти названия (в анонимном типе надо явно присваивать свойства, а у нас это не позволяет сделать разные наименования полей)
        // Присваивать названия надо после получения коллекции (collection) в методе
        public static void LoadFromCollection<TObject>(ExcelWorksheet sheet, (string literal, int numeric) address, TObject[] collection, string[] colums, bool printHeaders = false) where TObject : class
        {
            //Передать данные сразу из двух массивов?
            MemberInfo[] membersToInclude = typeof(TObject)
           .GetProperties(BindingFlags.Instance | BindingFlags.Public)
           .Where(p => colums.Contains(p.Name))
           .ToArray();


            //Загрузка данных коллекции по адресу

            // как добавить данные в существующую коллекцию в среде epplus ???



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

//Передаёт данные из массива в таблицу, а нам нужна замена существующих наименований полей


//public byte[] TestExcellGeneration_HorizontalLoadFromCollection()
//{
//    byte[] result = null;
//    using (ExcelPackage pck = new ExcelPackage())
//    {
//        var foo = pck.Workbook.Worksheets.Add("Foo");
//        var randomData = new[] { "Foo", "Bar", "Baz" }.ToList();
//        //foo.Cells["B4"].LoadFromCollection(randomData);

//        int startColumn = 2; // "B";
//        int startRow = 4;
//        for (int i = 0; i < randomData.Count; i++)
//        {
//            foo.Cells[startRow, startColumn + i].Value = randomData[i];
//        }

//        result = pck.GetAsByteArray();
//    }
//    return result;
//}

// Минусы: задаётся диапазон ячеек. При появлении нового столбца, надо будет вручную добавлять его в код.

//var exportedPersons = sheet.Cells["A2:E3"].ToCollectionWithMappings(row =>
//{
//    return new Person
//    {
//        FirstName = row.GetValue<string>("Fn"),
//        LastName = row.GetValue<string>("Ln"),
//        Height = row.GetValue<int>("H"),
//        BirthDate = row.GetValue<DateTime>("Bd")
//    };
//}, options => options.SetCustomHeaders("Fn", "Ln", "H", "Bd"));



//var membersToInclude = new List<MemberInfo>();
//var properties = typeof(TObject).GetProperties();
//foreach (var property in properties)
//{
//    if (colums.Contains(property.Name)) membersToInclude.Add(property);
//}