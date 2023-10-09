using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Obiecte
{
    //Segment=un segment folosit pt calcule introdus in formula la care cunoastem lungimea (sau aria) si nr de bucati
    public class Segment
    {
        private double valoare;
        private int cantitate;
        //poate fi m sau m2
        //private UM unitate_masura;

        public Segment(double valoare, int cantitate)
        {
            this.valoare = valoare;
            this.cantitate = cantitate;

        }

        public double Valoare   // property
        {
            get { return valoare; }   // get method
            set { valoare = value; }  // set method
        }

        public int Cantitate   // property
        {
            get { return cantitate; }   // get method
            set { cantitate = value; }  // set method
        }

        //public UM Unitate_masura   // property
        //{
        //    get { return unitate_masura; }   // get method
        //    set { unitate_masura = value; }  // set method
        //}
    }
}
