using System;
using NPOI.Utility.Excel;

namespace NPOI.Utility.Test
{
    public class Hotel
    {
        [Column(Title = "HotelName")]
        public string Name { get; set; }

        [Column(Title = "HoteAddress")]
        public string Address { get; set; }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            var hotels = ExcelFile.Read<Hotel>("D:\\Hotel_Info.xlsx", scheme =>
            {
                scheme.SheetIndex = 0;
                scheme.StartRow = 1;
            });
            Console.WriteLine("Hello World!");
        }
    }
}