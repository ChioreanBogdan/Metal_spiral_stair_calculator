using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Figura_geometrica
{
    //Contine calcule de arii
    abstract class Figura_geometrica
    {

        public abstract double Calculeaza_arie();
        //Calculeaza diametrul unui cerc folosind raza daca valoarea_introdusa="Raza" sau diametru daca valoarea_introdusa="Diametru" 

    }

    //aici am ramas 12-09-23
    class Cerc : Figura_geometrica
    {
        private double raza;
        private double diametru;
        public Cerc(double raz)
        {
            raza = raz;
            diametru = raz*2;
        }

        public override double Calculeaza_arie()
        {
            if (raza != null)
            {
                if (raza > 0)
                {
                    return 3.14 * raza * raza;
                }
                else return 0;
            }
            else return 0;
        }

        public double Calculeaza_circumferinta()
        {
            if (diametru != null)
            {
                if (diametru > 0)
                {
                    return 3.14 * diametru;
                }
                else return 0;
            }
            else return 0;
        }



        //public override double Calculeaza_arie(double val_raza_sau_diametru, string valoarea_introdusa)
        //{
        //    switch (valoarea_introdusa)
        //    {
        //        case "Diametru":

        //            if (val_raza_sau_diametru != null)
        //            {
        //                if (val_raza_sau_diametru > 0)
        //                {
        //                    return 3.14 * (val_raza_sau_diametru / 2 * (val_raza_sau_diametru / 2));
        //                }
        //                else return 0;
        //            }

        //            break;

        //        case "Raza":

        //            if (val_raza_sau_diametru != null)
        //            {
        //                if (val_raza_sau_diametru > 0)
        //                {
        //                    return 3.14 * (val_raza_sau_diametru * val_raza_sau_diametru);
        //                }
        //                else return 0;
        //            }

        //            break;
        //        default:

        //            return 0;

        //            break;
        //    }


        //    return 0;
        //}
    }

    class Patrulater : Figura_geometrica
    {
        private double lungime;
        private double latime;
        public Patrulater(double lung, double lat)
        {
            lungime = lung;
            latime = lat;
        }

        public override double Calculeaza_arie()
        {
            if ((lungime != null) && (latime != null))
            {
                if ((lungime > 0) && (latime > 0))
                {
                    return lungime*latime;
                }
                else return 0;
            }
            else return 0;
        }
    }
    class Patrat : Figura_geometrica
    {
        public Patrat(double lat)
        {
            latura = lat;
        }

        protected double latura { get; set; }

        public override double Calculeaza_arie()
        {
            if (latura != null)
            {
                if (latura > 0)
                {
                    return Math.Pow(latura, 2.0);
                }
                else return 0;
            }
            else return 0;
        }
    }

    //Oare trebuie sa pun si unghiurile in constructor
    class Triunghi_dreptunghic : Figura_geometrica
    {
        public Triunghi_dreptunghic(double cat1, double cat2)
        {
            cateta_1 = cat1;
            cateta_2 = cat2;
        }

        protected double cateta_1 { get; set; }
        protected double cateta_2 { get; set; }

        public override double Calculeaza_arie()
        {
            if ((cateta_1 != null) & (cateta_2 != null))
            {
                if ((cateta_1 > 0) & (cateta_2 > 0))
                {
                    return (cateta_1 * cateta_2) / 2;
                }
                else return 0;
            }
            else return 0;
        }
    }

    class Cub : Patrat
    {
        public Cub(double lat): base(lat)
        {
            latura = lat;
        }

        public double latura
        { 
            get
            {
                return latura;
            }
            set
            {
                latura = value;
            }
        }

        public override double Calculeaza_arie()
        {
            if (latura != null)
            {
                if (latura > 0)
                {
                    return 6*Math.Pow(latura, 2.0);
                }
                else return 0;
            }
            else return 0;
        }

        public double Calculeaza_volumul()
        {
            if (latura != null)
            {
                if (latura > 0)
                {
                    return Math.Pow(latura, 3.0);
                }
                else return 0;
            }
            else return 0;
        }
    }
}
