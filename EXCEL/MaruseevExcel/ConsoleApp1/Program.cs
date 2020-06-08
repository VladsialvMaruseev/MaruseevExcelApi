using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Exel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            _Application excel = new _Exel.Application();
            Workbook wb;
            Worksheet ws;
            int i = 0, j = 0;
            wb = excel.Workbooks.Open(@"D:\МПТ\Второй курс\2 семестр\c# Разработка кода ИС\Практические работы\Практическая работа №3\book1.xlsx");
            ws = wb.Worksheets[1];

            if (ws.Cells[i, j].Value2 != null)
                Console.WriteLine(ws.Cells[i, j].Value2);
            else
                Console.WriteLine("");
        }
    }
}
