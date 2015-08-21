using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinqToExcelNamedRangeTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = @"NamedRange.xlsb";
            var excel = new ExcelQueryFactory(file);

            var unique = excel.NamedRange("NamedRangeUniqueCell");
            var multiple = excel.NamedRange("NamedRangeMultipleCells");

        }
    }
}
