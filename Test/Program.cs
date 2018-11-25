using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\PC\Documents\Work task\ReffPriceV2\Test\bin\Debug\Test.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            xlRange.Cells[1, 1].Value2 = 1;
            xlRange.Cells[2, 1].Value2 = 2;
            xlRange.Cells[3, 1].Formula = "SUM(A1:A2)";

            xlWorkbook.Save();
            xlWorkbook.Close();

        }
    }
}
