using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Modules
{
    internal static class Operatiuni_Array
    {
        //Sterge toate aparitiile unui string dintr-un array sir_stringuri
        //apoi returneaza sir_stringuri fara string_de_eliminat
        public static string[] Sterge_string_din_array_string(string[] sir_stringuri,string string_de_eliminat)
        {
            sir_stringuri = sir_stringuri.Where(val => val != string_de_eliminat).ToArray();
            return sir_stringuri;
        }
    }
}
