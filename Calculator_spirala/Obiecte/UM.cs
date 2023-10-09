using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Obiecte
{
    //Aici am ramas 12-09-23
    //UM=scurt de la unitati de masura
    internal class UM
    {
        private string[] nume = { "m", "m2" };
        public string this[int i]
        {
            get { return nume[i]; }
            set { nume[i] = value; }
        }
    }
}
