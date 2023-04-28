using System;
using System.Collections.Generic;
using System.Linq;

namespace EpplusTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var productsDimensions = new List<ProductDimensionDTO>()
            {
                new ProductDimensionDTO
                {
                    Name = "product name 1",
                    Volume = 1,
                    Weight = 1,
                    AParam = "1",
                    BarCode = "barcode",
                    BParam = "1.00",
                    BpParam = "1.00",
                    BrandName = "GoodWill",
                    CParam = "1.00",
                    DParam = "1.00",
                    EParam = "1.00",
                    FormName = "formname",
                    FParam = "1.00",
                    GParam = "1.00",
                    HParam = "1.00",
                    NrParam = "1.00",
                    PackageCount = 10,
                    CategoryID = 8,
                    ProductID = 1,
                    ProductFormID = 1
                },
                new ProductDimensionDTO
                {
                    Name = "product name 2",
                    Volume = 1,
                    Weight = 1,
                    AParam = "1",
                    BarCode = "barcode",
                    BParam = "1.00",
                    BpParam = "1.00",
                    BrandName = "GoodWill",
                    CParam = "1.00",
                    DParam = "1.00",
                    EParam = "1.00",
                    FormName = "formname",
                    FParam = "1.00",
                    GParam = "1.00",
                    HParam = "1.00",
                    NrParam = "1.00",
                    PackageCount = 10,
                    CategoryID = 8,
                    ProductID = 2,
                    ProductFormID = 1
                },
                new ProductDimensionDTO
                {
                    Name = "product name 3",
                    Volume = 1,
                    Weight = 1,
                    AParam = "1",
                    BarCode = "barcode",
                    BParam = "1.00",
                    BpParam = "1.00",
                    BrandName = "GoodWill",
                    CParam = "1.00",
                    DParam = "1.00",
                    EParam = "1.00",
                    FormName = "formname",
                    FParam = "1.00",
                    GParam = "1.00",
                    HParam = "1.00",
                    NrParam = "1.00",
                    PackageCount = 10,
                    CategoryID = 8,
                    ProductID = 3,
                    ProductFormID = 1
                },
                new ProductDimensionDTO
                {
                    Name = "product name 4",
                    Volume = 1,
                    Weight = 1,
                    AParam = "1",
                    BarCode = "barcode",
                    BParam = "1.00",
                    BpParam = "1.00",
                    BrandName = "GoodWill",
                    CParam = "1.00",
                    DParam = "1.00",
                    EParam = "1.00",
                    FormName = "formname",
                    FParam = "1.00",
                    GParam = "1.00",
                    HParam = "1.00",
                    NrParam = "1.00",
                    PackageCount = 10,
                    CategoryID = 8,
                    ProductID = 4,
                    ProductFormID = 1
                },
                new ProductDimensionDTO
                {
                    Name = "product name 5",
                    Volume = 1,
                    Weight = 1,
                    AParam = "1",
                    BarCode = "barcode",
                    BParam = "1.00",
                    BpParam = "1.00",
                    BrandName = "GoodWill",
                    CParam = "1.00",
                    DParam = "1.00",
                    EParam = "1.00",
                    FormName = "formname",
                    FParam = "1.00",
                    GParam = "1.00",
                    HParam = "1.00",
                    NrParam = "1.00",
                    PackageCount = 10,
                    CategoryID = 8,
                    ProductID = 5,
                    ProductFormID = 1
                },
            };

            var filename = @"C:\Files\test.xlsx";
            var sheetname = "Лист 1";
            var columns = new[] { "Name", "AParam", "HParam" };
            var headers = new[] { "Высота", "Ширина", "Внутренний диаметр (B)" };

            ExportUtils.BaseExport<ProductDimensionDTO>(productsDimensions.ToArray(), ("A", 1), filename, sheetname, columns, headers);
        }
    }
}