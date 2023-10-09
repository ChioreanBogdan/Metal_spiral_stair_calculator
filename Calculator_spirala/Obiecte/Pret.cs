using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Calculator_spirala.Obiecte
{
    public class Pret
    {
        private double valoare_RON;

        private DateTime data_primire;

        public double Valoare_RON   // property
        {
            get { return valoare_RON; }   // get method
            set { valoare_RON = value; }  // set method
        }

        public DateTime Data_primire   // property
        {
            get { return data_primire; }   // get method
            set { data_primire = value; }  // set method
        }
    }
}
