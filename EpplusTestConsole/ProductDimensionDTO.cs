using System;

namespace EpplusTestConsole
{
    public class ProductDimensionDTO
    {
        public int ProductID { get; set; }
        public int CategoryID { get; set; }
        public int ProductFormID { get; set; }
        public string BarCode { get; set; }
        public string BrandName { get; set; }
        public string FormName { get; set; }
        public string Name { get; set; }
        public int PackageCount { get; set; }
        public double Volume { get; set; }
        public double Weight { get; set; }
        public string AParam { get; set; }
        public string BParam { get; set; }
        public string BpParam { get; set; }
        public string CParam { get; set; }
        public string DParam { get; set; }
        public string EParam { get; set; }
        public string FParam { get; set; }
        public string GParam { get; set; }
        public string HParam { get; set; }
        public string NrParam { get; set; }
        public decimal DecimalVolume => Math.Truncate((decimal)Volume * 10000000m) / 10000000m;
        public decimal DecimalWeight => Math.Truncate((decimal)Weight * 10000m) / 10000m;
    }
}