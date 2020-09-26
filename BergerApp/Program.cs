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
            //Console.WriteLine("Geben Sie die Länge des Informationsvektors ein.");
            // Die maximale Bitrate von Bergers klassischem Code.
            //double m = Convert.ToDouble(Console.ReadLine());
            double m = 3;
            // Das Arbeitsbuch.
            ExcelWorkbookGenerate(m);
        }

        static double ClassicBergerGenerate(double m)
        {
            double k = Math.Ceiling(Math.Log(m + 1, 2));
            //Console.WriteLine("Вид S(n,m)-кода Бергера: S({0},{1})", k + m, m);
            //Console.ReadLine();
            return k;
        }

        static void ExcelWorkbookGenerate(double m)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet GenerateBerger = workbook.CreateSheet("GenerateBerger");
            workbook.CreateSheet("Sheet A2");
            workbook.CreateSheet("Sheet A3");
            GenerateBerger.CreateRow(0).CreateCell(0).SetCellValue("Генерация заданного кода Бергера");
            GenerateBerger.AddMergedRegion(new CellRangeAddress(0, 0, 0, 3));
            // Tabelle mit Bergers klassischem Code.
            CreateTable(GenerateBerger, Convert.ToInt32(m));
            FileStream sw = File.Create("ErrorBerger.xlsx");
            workbook.Write(sw);
            sw.Close();
            
        }

        static void CreateTable(ISheet GenerateSheet, int m)
        {
            int k = Convert.ToInt32(ClassicBergerGenerate(m));
            IRow rowTitel = GenerateSheet.CreateRow(1);
            rowTitel.CreateCell(1).SetCellValue("Информационный вектор");
            rowTitel.CreateCell(1 + m).SetCellValue("Контрольный вектор");
            rowTitel.CreateCell(0).SetCellValue("№");
            GenerateSheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(1, 1, 1 + m, k + m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(1, 2, 0, 0));

            IRow rowVariabel = GenerateSheet.CreateRow(2);
            for (int i = 1; i <= (m + k); i++)
            {
                if (i <= m)
                {
                    string VariabelХ = "X" + (m - i).ToString();
                    rowVariabel.CreateCell(i).SetCellValue(VariabelХ);
                    continue;
                }
                string VariabelY = "Y" + (m + k - i).ToString();
                rowVariabel.CreateCell(i).SetCellValue(VariabelY);
            }
            
            int x = 1;
            for (int i = 1; i <= m; i++)
            {
                IRow row = GenerateSheet.CreateRow(i+2);
                for (int j = 0; j < 8; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
        }
    }
}
