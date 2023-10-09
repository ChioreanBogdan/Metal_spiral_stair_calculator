using Calculator_spirala.Obiecte;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.DataFormats;
using Excel = Microsoft.Office.Interop.Excel;

namespace Calculator_spirala.Modules
{
    internal static class Verificari
    {
        public static bool Verifica_daca_o_lista_int_nu_e_nula(List<int> lista_de_verificat)
        {
            if (lista_de_verificat != null)
            {

                return true;

            }
            else return false;
        }

        //verifica daca un workbook cu numele workbook_excel_de_verificat excel e deschis
        public static bool Verifica_daca_workbook_e_deschis(string workbook_excel_de_verificat)
        {
            try
            {
                Stream s = File.Open(workbook_excel_de_verificat, FileMode.Open, FileAccess.Read, FileShare.None);

                s.Close();

                //daca nu e deschis stream-ul se va putea deschide si inchide
                return false;
            }
            catch (Exception)
            {
                //daca deja e deschis vom avea exception error deci se va ajunge aci
                return true;
            }
        }

        //Verfica daca un string poate fi convertit in format DateTime
        public static bool Verifica_daca_string_e_data(string string_de_verificat)
        {
            DateTime dateTime;
            if (DateTime.TryParseExact(string_de_verificat, "dd-MM-yyyy", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out dateTime))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool Verifica_daca_string_e_data2(string string_de_verificat)
        {
            DateTime dateTime;
            if (DateTime.TryParse(string_de_verificat, out dateTime))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //Verifica daca un string string_de_verificat poate fi convertit in integer
        public static bool Verifica_daca_un_string_e_integer(string string_de_verificat)
        {
            bool e_integer = int.TryParse(string_de_verificat, out _);

            return e_integer;
        }

        //Verifica daca un string string_de_verificat poate fi convertit in double
        public static bool Verifica_daca_un_string_e_double(string string_de_verificat)
        {
            Double num = 0;
            bool e_Double = false;

            // Check for empty string.
            if (string.IsNullOrEmpty(string_de_verificat))
            {
                return false;
            }

            e_Double = Double.TryParse(string_de_verificat, out num);

            return e_Double;
        }

        //Verifica daca o lista de variabile int contine cel putin un element
        public static bool Verifica_daca_o_lista_de_preturi_nu_e_goala(List<Pret> lista_integer)
        {
            if (!(lista_integer == null))
            {
                if (lista_integer.Any())
                {
                    return true;
                }
            }
            return false;
        }

        //Verifica daca o lista de variabile int contine cel putin un element
        public static bool Verifica_daca_o_lista_int_nu_e_goala(List<int> lista_integer)
        {
            if (!(lista_integer==null))
            {
                if (lista_integer.Any())
                {
                    return true;
                }
            }
            return false;
        }
    }
}
