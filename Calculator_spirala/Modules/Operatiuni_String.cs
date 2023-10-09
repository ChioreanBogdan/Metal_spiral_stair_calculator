using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Calculator_spirala.Modules
{
    internal static class Operatiuni_String
    {
        //elimina toate aparitiile unui anumit caracter dintr-un string
        public static string Elimina_toate_substringurile_din_string(string string_de_modificat, string substring_de_eliminat)
        {
            string_de_modificat = string_de_modificat.Replace(substring_de_eliminat,"");

            return string_de_modificat;
        }

        //Elimina nr_de_caractere_de_eliminat dintr-un string string_de_modificat incepand de la pozitia pozitie_de_inceput 
        public static string Elimina_caractere_din_string(string string_de_modificat, int pozitie_de_inceput,int nr_de_caractere_de_eliminat)
        {
            if((pozitie_de_inceput>=0) & (nr_de_caractere_de_eliminat>0))
            {
                //Nu putem elimina pe o lungime>lungimea string-ului
                if(string_de_modificat.Length>(pozitie_de_inceput+ nr_de_caractere_de_eliminat))
                string_de_modificat=string_de_modificat.Remove(pozitie_de_inceput, nr_de_caractere_de_eliminat);

                return string_de_modificat;
            }
            return "";
        }

        //Adaug un substring substring_de_adaugat in string incepand de la pozitia pozitie_inceput
        public static string Adauga_un_substring_in_string(string string_de_modificat, string substring_de_adaugat,int pozitie_inceput)
        {
            if ((pozitie_inceput >= 0) & string_de_modificat.Length >= pozitie_inceput)
            {
                string_de_modificat = string_de_modificat.Insert(pozitie_inceput, substring_de_adaugat);

                return string_de_modificat;
            }
            return "";
        }

        //Verifica daca un string e double si returneaza tru,false daca nu e
        public static bool String_e_Double(string string_de_verificat)
        {
            if (double.TryParse(string_de_verificat, out double d) && !Double.IsNaN(d) && !Double.IsInfinity(d))
            {
                return true;
            }

            return false;
        }

        //Numara de cate ori apare un caracter caracter_de_numarat in string_de_analizat
        public static int De_cate_ori_apare_un_caracter_in_string(string string_de_analizat, char caracter_de_numarat)
        {
            return string_de_analizat.Count(t => t == caracter_de_numarat);
        }

        //Returneaza un nr de nr_caractere_de_returnat de la inceputul unui string
        public static string Returneaza_caractere_inceput_string(string string_de_analizat,int nr_caractere_de_returnat)
        {
            //verifica daca string_de_analizat nu e null si nr_caractere_de_returnat emai mare>0
            if (string.IsNullOrWhiteSpace(string_de_analizat) ==false && nr_caractere_de_returnat>0)
            {
                if (nr_caractere_de_returnat< string_de_analizat.Length)
                {
                    return string_de_analizat.Substring(0, nr_caractere_de_returnat);
                }
            }

            return "";
        }

        //Returneaza prima aparitie a unui caracter caracter_cautat in string_de_analizat
        public static int Returneaza_index_caracter_in_string(string string_de_analizat, char caracter_cautat)
        {
            if (string.IsNullOrWhiteSpace(string_de_analizat) == false && caracter_cautat != '\0')
            {
                if (string_de_analizat.ToLower().Contains(Char.ToString(caracter_cautat).ToLower()))
                {
                    return string_de_analizat.IndexOf(caracter_cautat);
                }
            }
            return 0;
        }

        //Returneaza string-ul dintre 2 char-uri (string_inceput si string_sfarsit)
        public static string Returneaza_string_intre_2_charuri(string string_de_analizat, char caracter_1, char caracter_2)
        {
            //verifica daca string_de_analizat,string_inceput,string_sfarsit nu e null 
            //'\0'- e null
            if (string.IsNullOrWhiteSpace(string_de_analizat) == false && caracter_1!= '\0' && caracter_2 != '\0')
            {

                //Daca string_de_analizat e <= 2 e imposibil de continuat
                if (string_de_analizat.Length>=2)
                {
                    Debug.WriteLine("Char.ToString(caracter_1)=" + Char.ToString(caracter_1));

                    if (string_de_analizat.ToLower().Contains(Char.ToString(caracter_1).ToLower()) && string_de_analizat.ToLower().Contains(Char.ToString(caracter_2).ToLower()))
                    {
                        int min = 0;
                        int max = 0;
                        //+1= lungimea caracterului
                        int index_c1 = string_de_analizat.IndexOf(caracter_1);

                        int index_c2 = string_de_analizat.IndexOf(caracter_2);

                        max=Altele.Max_2_int(index_c1, index_c2);
                        min = Altele.Min_2_int(index_c1, index_c2);
                        //min=min+1 pt ca altfel va fi inclus si primul caracter in String-ul rezultat
                        min = min + 1;

                        String rezultat = string_de_analizat.Substring(min, max-min);

                        return rezultat;
                    }
                        
                }
            }

            return "";
        }
    }
}
