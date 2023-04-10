using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EpplusTestConsole
{
    internal static class StylesUtils
    {
        public static void ProcessCommonStyles(this ExcelRangeBase range)
        {
            //range.Style.WrapText = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

        }

        public static void ProcessTypeFormat<T, F>(this ExcelRangeBase range, T[] collection, string format, bool printHeaders)
        {
            var properties = typeof(T).GetProperties();
            for (var i = 0; i < properties.Length; i++)
                if (properties[i].PropertyType == typeof(F))
                    range.Worksheet.Cells[range.Start.Row + (printHeaders ? 1 : 0), range.Start.Column + i, range.Start.Row + collection.Length, range.Start.Column + i]
                        .Style.Numberformat.Format = format;
        }

        public static void ProcessTypeStyles<T>(this ExcelRangeBase range, T[] collection, bool printHeaders)
        {
            ProcessTypeFormat<T, DateTime>(range, collection, "dd.MM.yyyy", printHeaders);
            ProcessTypeFormat<T, decimal>(range, collection, "#,##0.000;(#,##0.000)", printHeaders);
            ProcessTypeFormat<T, double>(range, collection, "#,##0.000000;(#,##0.000000)", printHeaders);
        }
    }
}