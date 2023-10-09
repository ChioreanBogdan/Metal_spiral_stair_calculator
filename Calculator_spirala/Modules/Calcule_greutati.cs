using Calculator_spirala.Figura_geometrica;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator_spirala.Modules
{
    internal static class Calcule_greutati
    {
        //returneaza greutatea  unei bare rotunde pt care se da diametrul,lungimea barei si densitatea materialului din care e facuta bara
        // Ø_bara se da in mm,lungime_bara se da in m,densitate_material se da in kg/m3
        public static double Calculeaza_greutate_bara_rotunda(double Ø_bara, double lungime_bara, double densitate_material)
        {
            double arie_cerc = 0;

            Ø_bara = Ø_bara / 1000;

            Cerc cerc_Ø_bara = new Cerc(Ø_bara / 2);
            arie_cerc = cerc_Ø_bara.Calculeaza_arie();

            if(arie_cerc>0)
            {
                return Math.Round(arie_cerc * lungime_bara * densitate_material,3);
            }

            return 0;
        }

        public static double Calculeaza_greutate_bara_dreptunghiulara(double L_bara, double l_bara, double lungime_bara, double densitate_material)
        {
            double arie_dreptunghi = 0;

            L_bara = L_bara / 1000;
            l_bara = l_bara / 1000;

            Patrulater dreptunghi_bara = new Patrulater(L_bara, l_bara);
            arie_dreptunghi = dreptunghi_bara.Calculeaza_arie();

            if (arie_dreptunghi > 0)
            {
                return Math.Round(arie_dreptunghi * lungime_bara * densitate_material, 3);
            }

            return 0;
        }

        //Genereaza o formula utilizabila in excel pt calculul unui material pt care cunoastem greutatea/m
        //densitate_material se da in kg/m
        public static string Genereaza_formula_cu_greutate_m_cunoscuta(string string_de_pus_in_paranteza, double greutate_per_metru)
        {
            return "=(" + string_de_pus_in_paranteza + ")*" + greutate_per_metru;
        }

        //Genereaza o formula utilizabila in excel pt calcululul greutatii unei table pt care cunoastem grosimea si densitatea materialului din care e facuta
        //grosime_tabla se da in mm
        //densitate_material se da in kg/m3
        public static string Genereaza_formula_greutate_tabla(string string_de_pus_in_paranteza, double grosime_tabla,double densitate_material)
        {
            grosime_tabla = grosime_tabla / 1000;
            return "=(" + string_de_pus_in_paranteza + ")*"+ grosime_tabla+"*"+ densitate_material;
        }
    }
}
