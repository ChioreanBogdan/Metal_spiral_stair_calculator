using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Modules
{
    //Aici am ramas 11-09-23
    internal static class Calcule_cantitati_brute
    {
        //Rotunjeste in sus lungime_de_rotunjit pana la o valoare divizibila cu lungime_fir,daca lungime_de_rotunjit<=(lungime_fir/2)
        //sau returneaza lungime_fir daca 
        public static double Rotunjeste_lungime_segment(double lungime_de_rotunjit, double lungime_fir)
        {
            double rest = 0;

            lungime_de_rotunjit = Math.Round(lungime_de_rotunjit, 2);
            //lungime_de_rotunjit = Math.Round(lungime_de_rotunjit,2);

            if (lungime_de_rotunjit<(lungime_fir/2))
            {
                do
                {
                    //% operator modulo (rest)
                    rest = Math.Round(lungime_fir % lungime_de_rotunjit, 2);
                    //Daca restul e zero nu mai are rost sa continuam
                    if((rest==0) | (Altele.Obtine_partea_decimala_din_double((lungime_fir / lungime_de_rotunjit), 2)==0))
                    {
                        break;
                    }

                    if (rest != 0)
                    {
                        lungime_de_rotunjit = lungime_de_rotunjit + 0.01;
                        lungime_de_rotunjit = Math.Round(lungime_de_rotunjit, 2);
                    }
                }
                while ((lungime_de_rotunjit <= (lungime_fir / 2)) | rest != 0);
                return Math.Round(lungime_de_rotunjit,2);
            }
            else if ((lungime_de_rotunjit > (lungime_fir / 2)) & (lungime_de_rotunjit<=lungime_fir))
            {
                return lungime_fir;
            }

            return 0;
        }

    }
}
