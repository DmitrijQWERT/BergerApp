using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BergerApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Geben Sie die Länge des Informationsvektors ein.");
            double m = Convert.ToDouble(Console.ReadLine());
            ClassicBergerGenerate(m);
            ExcelGenerate(m);
        }

        static void ClassicBergerGenerate(double m)
        {
            double k = Math.Ceiling(Math.Log(m + 1, 2));
            Console.WriteLine("Вид S(n,m)-кода Бергера: S({0},{1})", k + m, m);
            Console.ReadLine();
        }

        static void ExcelGenerate(double m)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet GenerateBerger = workbook.CreateSheet("GenerateBerger");
            GenerateBerger.CreateRow(0).CreateCell(0).SetCellValue("Генерация заданного кода Бергера");
            GenerateBerger.AddMergedRegion(new CellRangeAddress(0, 0, 0, 3));
            CreateTable(GenerateBerger, m);
            workbook.CreateSheet("Sheet A2");
            workbook.CreateSheet("Sheet A3");
            FileStream sw = File.Create("ErrorBerger.xlsx");
            workbook.Write(sw);
            sw.Close();
        }

        static void CreateTable(ISheet GenerateSheet, double m)
        {
            GenerateSheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 1 + Convert.ToInt32(m)));
            GenerateSheet.CreateRow(1).CreateCell(1).SetCellValue("Генерация заданного кода Бергера");
            //GenerateSheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 1 + Convert.ToInt32(m)));

            int x = 1;
            for (int i = 1; i <= m; i++)
            {
                IRow row = GenerateSheet.CreateRow(i);
                for (int j = 0; j < 8; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
        }
    }
}
