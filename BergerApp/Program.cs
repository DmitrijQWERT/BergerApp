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
            int m = 5;
            // Das Arbeitsbuch
            ExcelWorkbookGenerate(m);
        }

        //static double ClassicBergerGenerate(int m)
        //{
        //    double k = Math.Ceiling(Math.Log(m + 1, 2));
        //    //Console.WriteLine("Вид S(n,m)-кода Бергера: S({0},{1})", k + m, m);
        //    //Console.ReadLine();
        //    return k;
        //}

        static void ExcelWorkbookGenerate(int m)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet GenerateBerger = workbook.CreateSheet("GenerateBerger");
            workbook.CreateSheet("Sheet A2");
            workbook.CreateSheet("Sheet A3");
            GenerateBerger.CreateRow(0).CreateCell(0).SetCellValue("Генерация заданного кода Бергера");
            GenerateBerger.AddMergedRegion(new CellRangeAddress(0, 0, 0, 3));
            // Tabelle mit Bergers klassischem Code.
            int Schritt = 2;
            for (int i = m; i >= 2; i--)
            {
               Schritt = CreateTable(GenerateBerger, i, Schritt);
            }
            FileStream sw = File.Create("ErrorBerger.xlsx");
            workbook.Write(sw);
            sw.Close();
            
        }

        static int CreateTable(ISheet GenerateSheet, int m, int Schritt)
        {
            //int k = Convert.ToInt32(ClassicBergerGenerate(m));
            int k = Convert.ToInt32(Math.Ceiling(Math.Log(m + 1, 2)));
            IRow rowTitel = GenerateSheet.CreateRow(Schritt);
            IRow rowBerger = GenerateSheet.CreateRow(Schritt-1);
            string Berger = "Вид S(n,m)-кода Бергера: S("+(k+m).ToString()+","+(m).ToString()+").";
            rowBerger.CreateCell(1).SetCellValue(Berger);
            rowTitel.CreateCell(1).SetCellValue("Информационный вектор");
            rowTitel.CreateCell(1 + m).SetCellValue("Контрольный вектор");
            rowTitel.CreateCell(0).SetCellValue("№");
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt, 1, m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt, 1 + m, k + m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt+1, 0, 0));

            IRow rowVariabel = GenerateSheet.CreateRow(Schritt + 1);
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
            int NummerZeile = Convert.ToInt32(Math.Pow(2, m));
            for (int i = 0; i < NummerZeile ; i++)
            {
                IRow rowTabel = GenerateSheet.CreateRow(i + Schritt + 2);
                rowTabel.CreateCell(0).SetCellValue(i);
                int AusGangVariableX = 0;
                int AusGangVariableY = 0;
                int r = 0;
                for (int j = 1; j <= m; j++)
                {
                    AusGangVariableX = (i >> (m - j))&1;
                    r = AusGangVariableX == 1 ? (r+1) : (r + 0);
                    rowTabel.CreateCell(j).SetCellValue(AusGangVariableX);
                }
                for (int j = 1; j <= k; j++)
                {
                    AusGangVariableY = (r >> (k - j)) & 1;
                    rowTabel.CreateCell(m+j).SetCellValue(AusGangVariableY);
                }
            }

            return (NummerZeile + Schritt + 6);
        }
    }
}
