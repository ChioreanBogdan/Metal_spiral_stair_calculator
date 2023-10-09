using Calculator_spirala.Modules;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Obiecte
{
    //TODO:Ce ai scris aici o sa functioneze numa pt extragere date din sheet-urile M6 pana la M20
    //NU O SA FUNCTIONEZE PT sheet-ul Altele
    //nume ii un concept foarte vag,trebuie lamurit!
    abstract class Element_asamblare
    {
        private string nume;
        private double greutate;
        private double pret;

        public Element_asamblare(string nume, double greutate, double pret)
        {
            this.nume = nume;
            this.greutate = greutate;
            this.pret = pret;
        }

        public string Nume   // property
        {
            get { return nume; }   // get method
            set { nume = value; }  // set method
        }

        public double Greutate   // property
        {
            get { return greutate; }   // get method
            set { greutate = value; }  // set method
        }

        public double Pret   // property
        {
            get { return pret; }   // get method
            set { pret = value; }  // set method
        }
        public abstract double Extrage_greutate_din_tabel_suruburi(string nume,string adresa_fisier_suruburi);
        public abstract double Extrage_pret_din_tabel_suruburi(string nume, string adresa_fisier_suruburi);
    }

    class Surub : Element_asamblare
    {
        private int diametru_filet;
        private int lungime_filet; //lungime filetului surubului (fara cap)
        private string tip_cap; //hexagonal,inecat sau imbus
        //aici am ramas 18-09-23
        private string tip_finisaj; //Zn (zincat),TZn (Termozincat) sau A2 (inox)

        public Surub(string nume, double greutate, double pret, int diametru_filet,int lungime_filet, string tip_cap, string tip_finisaj) : base(nume, greutate, pret)
        {
            this.diametru_filet = diametru_filet;
            this.lungime_filet= lungime_filet;
            this.tip_cap = tip_cap;
            this.tip_finisaj = tip_finisaj;
        }

        public string Tip_cap   // property
        {
            get { return tip_cap; }   // get method
            set { tip_cap = value; }  // set method
        }

        //adresa fisierului de suruburi : "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"
        public override double Extrage_greutate_din_tabel_suruburi(string nume_surub, string adresa_fisier_suruburi)
        {
            //!=not
            //exemplu de nume surub: "Surub cap hexagonal"
            if (!string.IsNullOrEmpty(nume_surub))
            {
                int rand_tip_cap = 0;
                int coloana_tip_cap = 0;
                int rand_lungime_filet = 0;
                int coloana_masa = 0;

                double greutate_rezultata = 0;

                Operatiuni_Excel oe = new Operatiuni_Excel();

                bool exista_rand_tip_cap = int.TryParse(oe.Gaseste_adresa_text_in_sheet(nume_surub, adresa_fisier_suruburi, "M"+this.diametru_filet, "B1:Z100", "Rand"), out rand_tip_cap);
                bool exista_coloana_tip_cap = int.TryParse(oe.Gaseste_adresa_text_in_sheet(nume_surub, adresa_fisier_suruburi, "M" + this.diametru_filet, "B1:Z100", "Coloana"), out coloana_tip_cap);

                if ((rand_tip_cap > 0) & (coloana_tip_cap>0))
                {
                    string litera_coloana_tip_cap=oe.Returneaza_litera_coloana_dupa_nr(coloana_tip_cap);
                    //Cautam 
                    bool exista_rand_lungime_filet = int.TryParse(oe.Gaseste_adresa_text_in_sheet(this.lungime_filet.ToString(), adresa_fisier_suruburi, "M" + this.diametru_filet, litera_coloana_tip_cap+""+ rand_tip_cap + ":"+litera_coloana_tip_cap+"100", "Rand"), out rand_lungime_filet);
                    if (rand_lungime_filet>0)
                    {
                        bool exista_coloana_masa= int.TryParse(oe.Gaseste_adresa_text_in_sheet("Masa", adresa_fisier_suruburi, "M" + this.diametru_filet, litera_coloana_tip_cap + "" + rand_tip_cap + ":Z100", "Coloana"), out coloana_masa);
                        if (coloana_masa>0)
                        {
                            bool am_obtinut_o_greutate=double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_filet, coloana_masa, rand_lungime_filet),out greutate_rezultata);
                            if (am_obtinut_o_greutate==true)
                            {
                                return greutate_rezultata;
                            }
                        }                           
                    }
                }
            }

            return 0;
        }


        //ATENTIE: DACA AI IN TABELUL DE SURUBURI CELULE MERGE-UITE NU O SA LA GASEASCA!
        public override double Extrage_pret_din_tabel_suruburi(string nume_surub, string adresa_fisier_suruburi)
        {
            if (!string.IsNullOrEmpty(nume_surub))
            {
                int rand_tip_cap = 0;
                int coloana_tip_cap = 0;
                int rand_lungime_filet = 0;
                int coloana_pret = 0;

                double pret_rezultat = 0;

                Operatiuni_Excel oe = new Operatiuni_Excel();

                bool exista_rand_tip_cap = int.TryParse(oe.Gaseste_adresa_text_in_sheet(nume_surub, adresa_fisier_suruburi, "M" + this.diametru_filet, "B1:Z100", "Rand"), out rand_tip_cap);
                bool exista_coloana_tip_cap = int.TryParse(oe.Gaseste_adresa_text_in_sheet(nume_surub, adresa_fisier_suruburi, "M" + this.diametru_filet, "B1:Z100", "Coloana"), out coloana_tip_cap);

                if ((rand_tip_cap > 0) & (coloana_tip_cap > 0))
                {
                    string litera_coloana_tip_cap = oe.Returneaza_litera_coloana_dupa_nr(coloana_tip_cap);
                    //Cautam 
                    bool exista_rand_lungime_filet = int.TryParse(oe.Gaseste_adresa_text_in_sheet(this.lungime_filet.ToString(), adresa_fisier_suruburi, "M" + this.diametru_filet, litera_coloana_tip_cap + "" + rand_tip_cap + ":" + litera_coloana_tip_cap + "100", "Rand"), out rand_lungime_filet);
                    if (rand_lungime_filet > 0)
                    {
                        bool exista_coloana_pret = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Pret "+this.tip_finisaj, adresa_fisier_suruburi, "M" + this.diametru_filet, litera_coloana_tip_cap + "" + rand_tip_cap + ":Z100", "Coloana"), out coloana_pret);
                        if (coloana_pret > 0)
                        {
                            bool am_obtinut_un_pret = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_filet, coloana_pret, rand_lungime_filet), out pret_rezultat);
                            if (am_obtinut_un_pret == true)
                            {
                                return pret_rezultat;
                            }
                        }
                    }
                }
            }

            return 0;
        }
    }

    class Piulita : Element_asamblare
    {
        private int diametru_interior;
        private string tip; //hexagonala,buton
        private string tip_finisaj; //Zn (zincat),TZn (Termozincat) sau A2 (inox)

        public Piulita(string nume, double greutate, double pret, int diametru_interior, string tip, string tip_finisaj) : base(nume, greutate, pret)
        {
            this.diametru_interior = diametru_interior;
            this.tip = tip;
            this.tip_finisaj = tip_finisaj;
        }

        //adresa fisierului de suruburi : "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"
        public override double Extrage_greutate_din_tabel_suruburi(string nume,string adresa_fisier_suruburi)
        {
                int rand_piulite = 0;
                int coloana_piulite = 0;
                int coloana_masa = 0;

                double greutate_rezultata = 0;

                Operatiuni_Excel oe = new Operatiuni_Excel();

                //Verficam daca exista randul cu textul intitulat "Piulite hexagonale normale"
                bool exista_rand_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Piulite hexagonale normale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_piulite);
                bool exista_coloana_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Piulite hexagonale normale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_piulite);

                if ((rand_piulite > 0) & (coloana_piulite>0))
                {
                    string litera_coloana_piulite = oe.Returneaza_litera_coloana_dupa_nr(coloana_piulite);
                    //Cautam 
                        bool exista_coloana_masa = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Masa", adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_piulite + "" + rand_piulite + ":Z100", "Coloana"), out coloana_masa);
                        if (coloana_masa > 0)
                        {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Piulite hexagonale normale
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                            bool am_obtinut_o_greutate = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_masa, rand_piulite+3), out greutate_rezultata);
                            if (am_obtinut_o_greutate == true)
                            {
                                return greutate_rezultata;
                            }
                        }
                }
            return 0;
        }
        public override double Extrage_pret_din_tabel_suruburi(string nume, string adresa_fisier_suruburi)
        {
            int rand_piulite = 0;
            int coloana_piulite = 0;
            int coloana_finisaj = 0;

            double pret_rezultat = 0;

            Operatiuni_Excel oe = new Operatiuni_Excel();

            //Verficam daca exista randul cu textul intitulat "Piulite hexagonale normale"
            bool exista_rand_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Piulite hexagonale normale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_piulite);
            bool exista_coloana_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Piulite hexagonale normale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_piulite);

            if ((rand_piulite > 0) & (coloana_piulite > 0))
            {
                string litera_coloana_piulite = oe.Returneaza_litera_coloana_dupa_nr(coloana_piulite);
                //Cautam 
                bool exista_coloana_finisaj = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Pret "+this.tip_finisaj, adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_piulite + "" + rand_piulite + ":Z100", "Coloana"), out coloana_finisaj);
                if (coloana_finisaj > 0)
                {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Piulite hexagonale normale
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                    bool am_obtinut_un_pret = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_finisaj, rand_piulite + 3), out pret_rezultat);
                    if (am_obtinut_un_pret == true)
                    {
                        return pret_rezultat;
                    }
                }
            }
            return 0;
        }
    }

    class Saiba : Element_asamblare
    {
        private int diametru_interior;
        private string tip; //simpla,pt_profile_U_si_I
        private string tip_finisaj; //Zn (zincat),TZn (Termozincat) sau A2 (inox)

        public Saiba(string nume, double greutate, double pret, int diametru_interior, string tip, string tip_finisaj) : base(nume, greutate, pret)
        {
            this.diametru_interior = diametru_interior;
            this.tip = tip;
            this.tip_finisaj = tip_finisaj;
        }

        public override double Extrage_greutate_din_tabel_suruburi(string nume, string adresa_fisier_suruburi)
        {
            int rand_piulite = 0;
            int coloana_piulite = 0;
            int coloana_masa = 0;

            double greutate_rezultata = 0;

            Operatiuni_Excel oe = new Operatiuni_Excel();

            //Verficam daca exista randul cu textul intitulat "Piulite hexagonale normale"
            bool exista_rand_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe plate pt metale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_piulite);
            bool exista_coloana_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe plate pt metale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_piulite);

            if ((rand_piulite > 0) & (coloana_piulite > 0))
            {
                string litera_coloana_piulite = oe.Returneaza_litera_coloana_dupa_nr(coloana_piulite);
                //Cautam 
                bool exista_coloana_masa = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Masa", adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_piulite + "" + rand_piulite + ":Z100", "Coloana"), out coloana_masa);
                if (coloana_masa > 0)
                {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Piulite hexagonale normale
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                    bool am_obtinut_o_greutate = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_masa, rand_piulite + 3), out greutate_rezultata);
                    if (am_obtinut_o_greutate == true)
                    {
                        return greutate_rezultata;
                    }
                }
            }
            return 0;
        }
        public override double Extrage_pret_din_tabel_suruburi(string nume, string adresa_fisier_suruburi)
        {
            int rand_saibe = 0;
            int coloana_saibe = 0;
            int coloana_finisaj = 0;

            double pret_rezultat = 0;

            Operatiuni_Excel oe = new Operatiuni_Excel();

            //Verficam daca exista randul cu textul intitulat "Saibe plate pt metale"
            bool exista_rand_saibe = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe plate pt metale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_saibe);
            bool exista_coloana_saibe = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe plate pt metale", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_saibe);

            if ((rand_saibe > 0) & (coloana_saibe > 0))
            {
                string litera_coloana_saibe = oe.Returneaza_litera_coloana_dupa_nr(coloana_saibe);
                //Cautam 
                bool exista_coloana_finisaj = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Pret " + this.tip_finisaj, adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_saibe + "" + rand_saibe + ":Z100", "Coloana"), out coloana_finisaj);
                if (coloana_finisaj > 0)
                {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Saibe plate pt metale
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                    bool am_obtinut_un_pret = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_finisaj, rand_saibe + 3), out pret_rezultat);
                    if (am_obtinut_un_pret == true)
                    {
                        return pret_rezultat;
                    }
                }
            }
            return 0;
        }
    }

    class Saiba_Grower : Element_asamblare
    {
        private int diametru_interior;
        //private string tip; //simpla,pt_profile_U_si_I exista mai multe tipuri Grower?
        private string tip_finisaj; //Zn (zincat),TZn (Termozincat) sau A2 (inox)

        public Saiba_Grower(string nume, double greutate, double pret, int diametru_interior, string tip_finisaj) : base(nume, greutate, pret)
        {
            this.diametru_interior = diametru_interior;
            //this.tip = tip;
            this.tip_finisaj = tip_finisaj;
        }

        public override double Extrage_greutate_din_tabel_suruburi(string nume, string adresa_fisier_suruburi)
        {
            int rand_piulite = 0;
            int coloana_piulite = 0;
            int coloana_masa = 0;

            double greutate_rezultata = 0;

            Operatiuni_Excel oe = new Operatiuni_Excel();

            //Verficam daca exista randul cu textul intitulat "Piulite hexagonale normale"
            bool exista_rand_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe Grower MN", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_piulite);
            bool exista_coloana_piulite = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe Grower MN", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_piulite);

            if ((rand_piulite > 0) & (coloana_piulite > 0))
            {
                string litera_coloana_piulite = oe.Returneaza_litera_coloana_dupa_nr(coloana_piulite);
                //Cautam 
                bool exista_coloana_masa = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Masa", adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_piulite + "" + rand_piulite + ":Z100", "Coloana"), out coloana_masa);
                if (coloana_masa > 0)
                {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Piulite hexagonale normale
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                    bool am_obtinut_o_greutate = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_masa, rand_piulite + 3), out greutate_rezultata);
                    if (am_obtinut_o_greutate == true)
                    {
                        return greutate_rezultata;
                    }
                }
            }
            return 0;
        }
        public override double Extrage_pret_din_tabel_suruburi(string nume, string adresa_fisier_suruburi)
        {
            int rand_saibe = 0;
            int coloana_saibe = 0;
            int coloana_finisaj = 0;

            double pret_rezultat = 0;

            Operatiuni_Excel oe = new Operatiuni_Excel();

            //Verficam daca exista randul cu textul intitulat "Saibe plate pt metale"
            bool exista_rand_saibe = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe Grower MN", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Rand"), out rand_saibe);
            bool exista_coloana_saibe = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Saibe Grower MN", adresa_fisier_suruburi, "M" + this.diametru_interior, "B1:Z100", "Coloana"), out coloana_saibe);

            if ((rand_saibe > 0) & (coloana_saibe > 0))
            {
                string litera_coloana_saibe = oe.Returneaza_litera_coloana_dupa_nr(coloana_saibe);
                //Cautam 
                bool exista_coloana_finisaj = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Pret " + this.tip_finisaj, adresa_fisier_suruburi, "M" + this.diametru_interior, litera_coloana_saibe + "" + rand_saibe + ":Z100", "Coloana"), out coloana_finisaj);
                if (coloana_finisaj > 0)
                {
                    //O sa functioneze numai daca am un tabel de forma :
                    //Saibe Grower MN
                    //Masa
                    //[kg]
                    //Ceea ce nu ar trebui sa fie o problema pt ca toate de la M6 la M20 sunt asa
                    bool am_obtinut_un_pret = double.TryParse(oe.Returneaza_valoarea_de_la_adresa(adresa_fisier_suruburi, "M" + this.diametru_interior, coloana_finisaj, rand_saibe + 3), out pret_rezultat);
                    if (am_obtinut_un_pret == true)
                    {
                        return pret_rezultat;
                    }
                }
            }
            return 0;
        }
    }
}
