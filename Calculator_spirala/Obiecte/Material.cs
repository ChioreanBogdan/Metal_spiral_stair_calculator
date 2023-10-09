using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Obiecte
{
    abstract class Material
    {
        private string nume; // field
        private string unitate_masura; // field
        private double pret; //pretul materialului/unitatea de masura a materialului
        private string formula_excel_cantitate_neta; //formula compatibila excel pt calcul greutate cantitate neta
        private string formula_excel_cantitate_bruta; //formula compatibila excel pt calcul greutate cantitate bruta

        public Material(string nume, string unitate_masura, double pret)
        {
            this.nume = nume;
            this.unitate_masura = unitate_masura;
            this.pret = pret;

        }

        public string Nume   // property
        {
            get { return nume; }   // get method
            set { nume = value; }  // set method
        }

        public string Unitate_masura   // property
        {
            get { return unitate_masura; }   // get method
            set { unitate_masura = value; }  // set method
        }

        public double Pret   // property
        {
            get { return pret; }   // get method
            set { pret = value; }  // set method
        }

        //Nu o punem in constructor
        public string Formula_excel_cantitate_neta   // property
        {
            get { return formula_excel_cantitate_neta; }   // get method
            set { formula_excel_cantitate_neta = value; }  // set method
        }

        //Nu o punem in constructor
        public string Formula_excel_cantitate_bruta   // property
        {
            get { return formula_excel_cantitate_bruta; }   // get method
            set { formula_excel_cantitate_bruta = value; }  // set method
        }

        //String pt ca vrem rezultatul returnat sub forma de formula pe care sa il inseram in excel
        public abstract string Calculeaza_greutatea(string sir_lungimi);
        public abstract string Calculeaza_suprafata(string sir_lungimi);
        public abstract string Calculeaza_cantitatea_bruta(string sir_lungimi);
    }

    class Teava_rotunda : Material  // derived class (child)
    {
        public double diametru;
        public double grosime;
        public double greutate_specifica; //kg/m

        public double Diametru   // property
        {
            get { return diametru; }   // get method
            set { diametru = value; }  // set method
        }

        public double Grosime   // property
        {
            get { return grosime; }   // get method
            set { grosime = value; }  // set method
        }

        public double Greutate_specifica   // property
        {
            get { return greutate_specifica; }   // get method
            set { greutate_specifica = value; }  // set method
        }

        public Teava_rotunda(string nume, string unitate_masura, double pret, double diametru, double grosime, double greutate_specifica) : base(nume, unitate_masura, pret)
        {
            this.diametru = diametru;
            this.grosime = grosime;
            this.greutate_specifica = greutate_specifica;
        }

        public override string Calculeaza_greutatea(string sir_lungimi)
        {
            string formula_rezultata = "";

            //!=not
            if (!string.IsNullOrEmpty(sir_lungimi))
            {
                //Sir_lungimi trebuie sa fie sub forma "1+2+3.2+9" cand e pass-uit
                formula_rezultata = "=(" + sir_lungimi + ")*" + Greutate_specifica;
            }

            return formula_rezultata;
        }

        //Aici am ramas 06-09-23
        public override string Calculeaza_suprafata(string sir_lungimi)
        {
            string formula_rezultata = "";
            double circumferinta_teava=0;
            double Diametru_gaura = 0;

            double Raza = 0;
            double Raza_gaura = 0;
            //=aria calculata cu diametrul exterior
            double arie_exterior = 0;
            //=aria gaurii tejii
            double arie_gaura = 0;
            //=arie_exterior-aria gaurii
            double arie_capat = 0;

            Figura_geometrica.Cerc c_ext=new Figura_geometrica.Cerc(diametru/2);

            circumferinta_teava = c_ext.Calculeaza_circumferinta();

            arie_exterior = c_ext.Calculeaza_arie();

            Figura_geometrica.Cerc c_gaura = new Figura_geometrica.Cerc((diametru-grosime*2) / 2);

            arie_gaura = c_gaura.Calculeaza_arie();

            arie_capat = arie_exterior - arie_gaura;

            //!=not
            if (!string.IsNullOrEmpty(sir_lungimi) & circumferinta_teava > 0)
            {
                //Sir_lungimi trebuie sa fie sub forma "1+2+3.2+9" cand e pass-uit
                formula_rezultata = "=(" + sir_lungimi + ")*" + (circumferinta_teava/1000);
            }

            return formula_rezultata;
        }
        public override string Calculeaza_cantitatea_bruta(string sir_lungimi)
        {
            string formula_rezultata = "";

            return formula_rezultata;
        }

    }
}
