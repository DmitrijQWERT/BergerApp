using System;
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
            double m = 0;
            Console.WriteLine("Geben Sie die Länge des Informationsvektors ein.");
            m = Convert.ToDouble(Console.ReadLine());
            ClassicBergerGenerate(m);
        }

        static void ClassicBergerGenerate(double m)
        {
            double k = 0;
            k = Math.Ceiling(Math.Log(m+1,2));
            Console.WriteLine("Вид S(n,m)-кода Бергера: S({0},{1})", k + m, m);
            Console.ReadLine();
        }
    }
}
