using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Modules
{
    internal static class Altele
    {
        //Returneaza maximul intre 2 integeri
        public static int Max_2_int(int num1, int num2)
        {
            if((num1!=null) && (num2 != null))
            {
                    if (num1>num2)
                    {
                        return num1;
                    }
                    else
                    {
                        return num2;
                    }
            }
            return 0;
        }

        public static int Min_2_int(int num1, int num2)
        {
            if ((num1 != null) && (num2 != null))
            {
                if (num1 > num2)
                {
                    return num2;
                }
                else
                {
                    return num1;
                }
            }
            return 0;
        }

        //Obtine partea decimala dintr-un double returnand nr_de_decimale_de_obtinut decimale
        public static double Obtine_partea_decimala_din_double(double nr_de_analizat,int nr_de_decimale_de_obtinut)
        {
            double partea_decimala;
            //var floatNumber = nr_de_analizat;

            //var partea_decimala = floatNumber - Math.Truncate(floatNumber);

            Debug.WriteLine("Math.Truncate(nr_de_analizat)="+ Math.Truncate(nr_de_analizat));

            partea_decimala = nr_de_analizat - Math.Truncate(nr_de_analizat);

            return Math.Round(partea_decimala, nr_de_decimale_de_obtinut);
        }

    }
}
