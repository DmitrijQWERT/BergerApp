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
            int m = 3;
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
            IRow rowVariabel = GenerateSheet.CreateRow(Schritt + 1);
            string Berger = "Вид S(n,m)-кода Бергера: S("+(k+m).ToString()+","+(m).ToString()+").";
            rowBerger.CreateCell(1).SetCellValue(Berger);
            rowBerger.CreateCell(m + k + 1).SetCellValue("Анализ кода на обнаружение ошибок. Необнаруживаемые ошибки кратности d:");
            rowTitel.CreateCell(1).SetCellValue("Информационный вектор");
            rowTitel.CreateCell(1 + m).SetCellValue("Контрольный вектор");
            rowTitel.CreateCell(0).SetCellValue("№");
            for (int i = 1; i <= m; i++)
            {
                rowVariabel.CreateCell(m + k + i).SetCellValue(i);
            }
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt, 1, m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt, 1 + m, k + m));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt, Schritt + 1, 0, 0));
            GenerateSheet.AddMergedRegion(new CellRangeAddress(Schritt - 1, Schritt, m + k + 1, m + k + 10));
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
                Fehleranalyse(GenerateSheet, rowTitel, rowBerger, rowTabel, Schritt, m, k, r, i);
            }
            return (NummerZeile + Schritt + 6);
        }
        static void Fehleranalyse(ISheet GenerateSheet, IRow rowTitel, IRow rowBerger, IRow rowTabel, int Schritt, int m, int k, int r, int Kombination)
        {
            int AusGangVariableX = 0;
            int EntstellungKombination = 0;
            for (int i = 1; i <= m; i++) // кратность ошибки
            {
                int SchrittX0 = Convert.ToInt32(((Kombination & 1) << m)| Kombination);
                int Maske = Convert.ToInt32(Math.Pow(2, i)) - 1;
                int EinsMaske = Convert.ToInt32(Math.Pow(2, m)) - 1;
                int EinsMaskeX0 = Convert.ToInt32(Math.Pow(2, m)) - 2;
                string EntstellungString = " ";
                int rEntst = 0;
                int Entstellung = 0;
                for (int j = 1; j <= m; j++) // поэлементная проверка
                {
                    if ((i == 1) || j != 1)
                    {
                        EntstellungKombination = 0;
                        EntstellungKombination = Kombination ^ (Maske << m - j);
                        rEntst = 0;
                        for (int l = 1; l <= m; l++)//Определение веса
                        {
                            AusGangVariableX = (EntstellungKombination >> (m - l)) & 1;
                            rEntst = AusGangVariableX == 1 ? (rEntst + 1) : (rEntst + 0);
                        }
                    }
                    else if (i !=1 & j == 1)
                    {
                        EntstellungKombination = 0;
                        EntstellungKombination = SchrittX0 ^ (Maske << m - j);
                        EntstellungKombination = ((EntstellungKombination >> m) | ((EinsMaske & EntstellungKombination) & EinsMaskeX0));
                        rEntst = 0;
                        for (int l = 1; l <= m; l++)//Определение веса
                        {
                            AusGangVariableX = (EntstellungKombination >> (m - l)) & 1;
                            rEntst = AusGangVariableX == 1 ? (rEntst + 1) : (rEntst + 0);
                        }                       
                    }
                    if (rEntst == r)
                    {
                        if ((m == 2) & (j == 2))
                        {
                            continue;
                        }
                        EntstellungString += Convert.ToString(EntstellungKombination, 2) + " ";
                        Entstellung++;
                    }
                }
                EntstellungString += Entstellung + " ";
                rowTabel.CreateCell(m + k + i).SetCellValue(EntstellungString);
            }   
        }
    }
}
