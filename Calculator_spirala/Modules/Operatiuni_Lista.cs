using Calculator_spirala.Obiecte;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Modules
{
    internal class Operatiuni_Lista
    {
        //Creaza o lista noua de obiecte de tip segment
        public List<Segment> Creaza_lista_de_segmente_noua()
        {
            List<Segment> lista_segmente_noua = new List<Segment>();

            return lista_segmente_noua;
        }
        
        //Adauga un obiect de tip "Segment" intr-o lista de segmente si returneaza lista cu elementul adaugat
        public List<Segment> Adauga_un_segment_in_lista_de_segmente(List<Segment> lista_segmente, Segment segment_de_adaugat)
        {
            lista_segmente.Add(segment_de_adaugat);

            return lista_segmente;
        }

        //Transforma o lista de obiecte de tip "Segment" intr-un string de ex: "1.2*3+1.5*2+3*1"
        public string Transforma_lista_segmente_in_string(List<Segment> lista_segmente)
        {
            string string_de_returnat = "";

            for(int i=0;i< lista_segmente.Count;i++)
            {
            Console.WriteLine("Valoarea "+i+" a segmentului: "+lista_segmente[i].Valoare);
            Console.WriteLine("Cantitatea " + i + " a segmentului: " + lista_segmente[i].Cantitate);
                //Daca elementul cu nr i nu e ultimul element din lista punem un "+" in coada
                if (i < (lista_segmente.Count-1))
                {
                    string_de_returnat = string_de_returnat + lista_segmente[i].Valoare.ToString() + "*" + lista_segmente[i].Cantitate.ToString() + "+";
                }
                //Daca elementul cu nr i nu e ultimul element din lista nu punem un "+" in coada
                else if (i == (lista_segmente.Count-1))
                {
                    string_de_returnat = string_de_returnat + lista_segmente[i].Valoare.ToString() + "*" + lista_segmente[i].Cantitate.ToString();
                }
            }

            return string_de_returnat;
        }


    }
}
