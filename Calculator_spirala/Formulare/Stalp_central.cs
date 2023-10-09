using Calculator_spirala.Modules;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

using System.Diagnostics;
using Calculator_spirala.Obiecte;
using Microsoft.Office.Interop.Word;
using System.Diagnostics.Eventing.Reader;
using Calculator_spirala.Figura_geometrica;

namespace Calculator_spirala
{
    public partial class Stalp_central : Form
    {
        public Stalp_central()
        {
            InitializeComponent();

            //Prima data presupunem ca stalpul nu are capac
            checkBox_capac.Checked = false;
            Poza_capac_stalp.Image = Properties.Resources.Nu_are_capac_protectie;

            //blocam butonul "Calculeaza" pana cand se completeaza corect datele din form
            //buton_calc_stalp.Enabled = false;

            //Sa nu fie ComboBox-urile goale la initializare
            if (comboBox_fi_teava.Items.Count > 0) { comboBox_fi_teava.SelectedItem = comboBox_fi_teava.Items[0]; }
            if (comboBox_fi_teava_sus.Items.Count > 0) { comboBox_fi_teava_sus.SelectedItem = comboBox_fi_teava_sus.Items[0]; }
            if (comboBox_gr_intreruperi.Items.Count > 0) { comboBox_gr_intreruperi.SelectedItem = comboBox_gr_intreruperi.Items[0]; }
            if (comboBox_gr_talpa.Items.Count > 0) { comboBox_gr_talpa.SelectedItem = comboBox_gr_talpa.Items[0]; }
            if (comboBox_guseu_intreruperi.Items.Count > 0) { comboBox_guseu_intreruperi.SelectedItem = comboBox_guseu_intreruperi.Items[0]; }
            if (comboBox_guseu_talpa.Items.Count > 0) { comboBox_guseu_talpa.SelectedItem = comboBox_guseu_talpa.Items[0]; }
        }

        //cand bifez/debifez Checkbox-ul "Are capac"
        private void checkBox_capac_CheckStateChanged(object sender, EventArgs e)
        {

            //MessageBox.Show("Semafoarele e pisici");
            if (checkBox_capac.Checked)
            {
                Poza_capac_stalp.Image = Properties.Resources.Are_capac_protectie;
            }
            else
            {
                Poza_capac_stalp.Image = Properties.Resources.Nu_are_capac_protectie;
            }
        }

        //Se pot introduce doar caractere numerice si doar un singur separator decimal
        private void textBox_lungime_fir_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        //Cand dam click in alta parte de pe form,vrem sa se stearga tot ce e in TextBox daca e <=0
        private void textBox_lungime_fir_Leave(object sender, EventArgs e)
        {
            //Operatiuni_String os = new Operatiuni_String();

            //MessageBox.Show("0 apare de: " + os.De_cate_ori_apare_un_caracter_in_string(textBox_lungime_fir.Text, '0'));
            //MessageBox.Show("Textul din textBox_lungime_fir e de: " + textBox_lungime_fir.Text.Length);

            if ((Operatiuni_String.De_cate_ori_apare_un_caracter_in_string(textBox_lungime_fir.Text, '0') == textBox_lungime_fir.Text.Length) | ((Operatiuni_String.De_cate_ori_apare_un_caracter_in_string(textBox_lungime_fir.Text, '0') + Operatiuni_String.De_cate_ori_apare_un_caracter_in_string(textBox_lungime_fir.Text, '.')) == textBox_lungime_fir.Text.Length))
            {
                textBox_lungime_fir.Text = "";
            }

        }

        //Butonul "Calculeaza" de pe form-ul "Stalp_central"

        private void buton_calc_stalp_Click(object sender, EventArgs e)
        {
            int coloana_test;
            int rand_test;

            Debug.WriteLine("0.18 rotunjit pt fir de 6 e: "+ Calcule_cantitati_brute.Rotunjeste_lungime_segment(0.18,6));

            //Teava din care se taie stalpul central,poate fi Ø146x10 sau Ø127x8
            Teava_rotunda teava_stalp_central = new Teava_rotunda("", "", 0, 0, 0, 0);
            //Teava din care se taie bucsile pt stalpul central (Ø127x8  in cazul in care e Ø146x10 si ??? pt Ø127x8)
            Teava_rotunda teava_bucsa = new Teava_rotunda("", "", 0, 0, 0, 0);
            //Teava din care se realizaeaza teava de sus (poate fi Ø48x3 sau Ø60.3x3) ?       
            Teava_rotunda teava_sus = new Teava_rotunda("", "", 0, 0, 0, 0);

            //Numarul de materiale (teji,table,suruburi,etc) necesare pt a fabrica stalpul
            //Incepem cu un minim de 8 materiale:
            //aici am ramas la 02-10-23
            // - teava pt stalpul central
            // - teava sus
            // - capac pt teava sus (tabla 3mm)
            // - tabla pt flanse+guseuri
            //-surub pt prindere flansa cu teava sus de flansa cu stalpul central
            //Trebuie lamurit!
            //-surub pt prindere flansa cu teava sus de flansa cu stalpul central (partea cu tabla de 20mm)
            //-piulita pt prindere flansa cu teava sus de flansa cu stalpul central
            //-saiba plata pt prindere flansa cu teava sus de flansa cu stalpul central
            //-saiba Grower pt prindere flansa cu teava sus de flansa cu stalpul central
            int numar_materiale = 8;

            //Variabile in care vom memora datele din formularul Stalp_central (daca au fost introduse corect)
            //Lungimea firului din care se taie stalpul/segmentii de stalp pt coloana
            double lungime_fir_coloana = 0;
            //lungimea de la talpa pana deasupra flansei de sus a coloanei centrale a spiralei 
            double lungime_totala_coloana = 0;
            //segment coloana=lungime_totala_coloana daca nu avem intreruperi sau lungime_segment_coloana =lungime_totala_coloana/nr_intreruperi daca nr_intreruperi>0
            double lungime_segment_coloana = 0;
            //nr_segmente_coloana=nr_intreruperi+1
            int nr_segmente_coloana = 0;
            //Formula segment coloana: Va fi "lungime_segment_coloana*nr_segmente_coloana" daca avem intreruperi sau
            //"lungime_segment_coloana" pt un singur segment

            //Ø-ul talpii de jos si grosimea
            double Ø_flansa_talpa = 0;

            //O lista cu toate preturile pt grosimea selectata,ordonata de la cel mai recent (prima pozitie) la cel mai vechi (ultima pozitie) 
            List<Pret> lista_preturi_tabla_talpa = new List<Pret>();
            double grosime_talpa = 0;
            double pret_tabla_talpa = 0;

            int nr_intreruperi = 0;
            int nr_flanse = 0;
            //Ø-ul flansei de intrerupere (daca exista) si a celei de sus
            double Ø_flansa = 0;
            double grosime_flansa_intrerupere = 0;

            //lungimea tevii de sus
            double lungime_teava_sus = 0;

            //Aici memoram nr coloanei pe care se afla UM,greutatea specifica si Ultimul pret in sheet-ul "12.Teava rotunda(Round Pipe)" pt teava din care e facuta coloana centrala
            int nr_coloana_UM_tevi = 0;
            int nr_coloana_greutate_tevi = 0;
            int nr_coloana_pret_tevi = 0;

            //Aici memoram nr coloanei pe care se afla UM,greutatea specifica si Ultimul pret in sheet-ul "8.Tabla neteda"
            int nr_coloana_UM_tabla_neteda = 0;
            int nr_coloana_greutate_tabla_neteda = 0;
            int nr_coloana_pret_tabla_neteda = 0;

            //Aici memoram randul din "PRETURI MATERIALE" pe care gasim teava din care e facut stalpul central
            int nr_rand_teava_centrala = 0;
            //Aici memoram randul din "PRETURI MATERIALE" pe care gasim teava din care e facuta teava de sus
            int nr_rand_teava_sus = 0;
            //Aici memoram randul din "PRETURI MATERIALE" pe care gasim teava din care e facuta teava pt bucsa
            int nr_rand_teava_bucse = 0;

            //Formula tabla 12/10/8 se va realiza la final si se va schimba in functie
            string formula_tabla_8mm = "";
            string formula_tabla_10mm = "";
            string formula_tabla_12mm = "";

            //Bool ce devine true cand am toate datele necesare pt a calulca un stalp central (sunt completate toate campurile necesare cu date valide)
            bool Stalpul_se_poate_calcula = false;

            //Adresa sub forma unui string a fisierului excel unde vom pune rezultatul pe baza datelor introduse in form
            string adresa_template_spirala= "S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls";
            string sheet_template_spirala = "Spirala";

            //Adresa sub forma unui string a fisierului excel unde tinem date despre pretul,greutatea.etc a materialelor folosite pt spirala
            string adresa_preturi_materiale = "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm";

            Operatiuni_Excel oe = new Operatiuni_Excel();
            //oe.Gaseste_informatii_text_in_sector_sheet("Pillar","Raw material",adresa_template_spirala, sheet_template_spirala,"A1:G100","Rand");
            //rand_inceput_pillar=+1 la randul din worksheet unde apare scris "Pillar"
            int rand_inceput_pillar =Int32.Parse(oe.Gaseste_informatii_text_in_sector_sheet("Pillar", "Raw material", adresa_template_spirala, sheet_template_spirala, "A1:G100", "Rand"));
            int rand_sfarsit_pillar= Int32.Parse(oe.Gaseste_informatii_text_in_sector_sheet("Repos", "Raw material", adresa_template_spirala, sheet_template_spirala, "A1:G100", "Rand"));

            //Stergem prima celula sub Denumirea "Pillar"

            //rand_inceput_pillar+1 si rand_sfarsit_pillar-1 pt a nu goli si continutul randurilor cu "Pillar" si "Repos"
            oe.Sterge_continut_range(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar+1, 1, rand_sfarsit_pillar-1, 7);
            int nr_randuri_de_sters = (rand_sfarsit_pillar - 1) - (rand_inceput_pillar + 1);
            //Daca stergem tot ce e intre rubricile cu "Pillar" si "Repos" stergem si formatarea,de aceea trebuie sa lasam un rand si sa incepem de la rand_inceput_pillar + 2
            oe.Sterge_randuri_worksheet(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, nr_randuri_de_sters);
            //+2 pt ca randul cu formatarile sa ramana primul
            oe.Insereaza_randuri_worksheet(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 9);
            oe.Adauga_chenare_pt_celula_in_range(adresa_template_spirala, sheet_template_spirala,rand_inceput_pillar + 2,1, rand_inceput_pillar + 2+8,7);
            //Operatiuni_String os = new Operatiuni_String();
            //oe.Deschide_preturi();
            //oe.Creaza_workbook_excel_nou("Tabel spirala noua");

            int rand_stalp = 0;
            int rand_repos = 0;

            //oe.Gaseste_adresa_celula_ce_incepe_cu_stringul

            //Verific daca textBox_nr_Intreruperi,textBox_lungime_totala si textBox_lungime_totala au valori numerice
            bool nr_intreruperi_e_numeric = int.TryParse(textBox_nr_Intreruperi.Text, out nr_intreruperi);
            bool lungime_totala_coloana_e_numeric = double.TryParse(textBox_lungime_totala.Text, out lungime_totala_coloana);
            bool lungime_fir_coloana_e_numeric = double.TryParse(textBox_lungime_fir.Text, out lungime_fir_coloana);
            bool lungime_teava_sus_e_numeric = double.TryParse(textBox_L_teava_sus.Text, out lungime_teava_sus);

            //Verfic adaca textbox textBox_fi_talpa contine o valoare numerica
            bool Ø_flansa_talpa_e_numeric = double.TryParse(textBox_fi_talpa.Text, out Ø_flansa_talpa);
            //Ø_flansa=flansa folosita pt flansa care conecteaza teava de sus de stalpul central si pt intreruperi
            bool Ø_flansa_e_numeric = double.TryParse(textBox_fi_intreruperi.Text, out Ø_flansa);

            //aici am ramas 05-10-23
            MessageBox.Show("Form-ul e completat bine:" + Verifica_daca_formul_e_completat_bine(nr_intreruperi, lungime_totala_coloana, lungime_fir_coloana));

            MessageBox.Show("Nr grosimi tabla folosite:" + Cate_grosimi_de_tabla_avem());

            if (Ø_flansa_talpa_e_numeric == true)
            {
                double arie_Ø_flansa_talpa = new Cerc(Ø_flansa_talpa / 1000 / 2).Calculeaza_arie();
                arie_Ø_flansa_talpa = Math.Round(arie_Ø_flansa_talpa, 3);

                //Daca e selectat 10mm adaugam aria flansei de jos la formula_tabla_10mm
                if (comboBox_gr_talpa.Text == "12mm")
                {
                    formula_tabla_12mm = formula_tabla_12mm + arie_Ø_flansa_talpa + "+";
                }
                else if (comboBox_gr_talpa.Text == "10mm")
                {
                    formula_tabla_10mm = formula_tabla_10mm + arie_Ø_flansa_talpa + "+";
                }
                else if (comboBox_gr_talpa.Text == "8mm")
                {
                    formula_tabla_8mm = formula_tabla_8mm + arie_Ø_flansa_talpa + "+";
                }

                //MessageBox.Show("Aria guseului de la talpa e:" + Returneaza_arie_guseu_talpa(comboBox_fi_teava.Text, Ø_flansa_talpa));

                double arie_guseu_talpa = Returneaza_arie_guseu_talpa(comboBox_fi_teava.Text, Ø_flansa_talpa);

                if (arie_guseu_talpa > 0)
                {
                    //Punem intariturile (guseurile) de la talpa
                    if (comboBox_guseu_talpa.Text == "12mm")
                    {
                        formula_tabla_12mm = formula_tabla_12mm + arie_guseu_talpa + "*4+";
                    }
                    else if (comboBox_guseu_talpa.Text == "10mm")
                    {
                        formula_tabla_10mm = formula_tabla_10mm + arie_guseu_talpa + "*4+";
                    }
                    else if (comboBox_guseu_talpa.Text == "8mm")
                    {
                        formula_tabla_8mm = formula_tabla_8mm + arie_guseu_talpa + "*4+";
                    }
                }

                //MessageBox.Show("Formula cantitate neta talpa: "+Calcule_greutati.Genereaza_formula_tabla(arie_Ø_flansa_talpa.ToString(),10,7850));
            }

            if (nr_intreruperi_e_numeric == true)
            {
                //cate 2 flanse/intrerupere + 2 flanse sus
                nr_flanse = nr_intreruperi * 2 + 2;
                nr_segmente_coloana = nr_intreruperi + 1;
            }
            else
            {
                //2 flanse sus
                nr_flanse = 2;
                //Daca nu e completata rubrica "Nr intreruperi" sau e 0 -> ca nu avem intreruperi si avem un singur segment
                nr_segmente_coloana = 1;
            }

            if (Ø_flansa_e_numeric == true)
            {
                double arie_Ø_flansa = new Cerc(Ø_flansa / 1000 / 2).Calculeaza_arie();
                arie_Ø_flansa = Math.Round(arie_Ø_flansa, 3);

                MessageBox.Show("Aria guseului de la flansa e:" + Returneaza_arie_guseu_flansa(comboBox_fi_teava.Text, Ø_flansa));

                if (comboBox_gr_talpa.Text == "12mm")
                {
                    formula_tabla_12mm = formula_tabla_12mm + arie_Ø_flansa + "*" + nr_flanse + "+";
                }

                //Daca e selectat 10mm adaugam aria flansei de jos la formula_tabla_10mm
                if (comboBox_gr_talpa.Text == "10mm")
                {
                    formula_tabla_10mm = formula_tabla_10mm + arie_Ø_flansa + "*" + nr_flanse + "+";
                }
                else if (comboBox_gr_talpa.Text == "8mm")
                {
                    formula_tabla_8mm = formula_tabla_8mm + arie_Ø_flansa + "*" + nr_flanse + "+";
                }

                //Cate 3 bucati pt flansa de sus si 6 pt celelalte intreruperi (3 sus+3 jos)
                double arie_guseu_flansa = Returneaza_arie_guseu_flansa(comboBox_fi_teava.Text, Ø_flansa);

                if (arie_guseu_flansa > 0)
                {
                    //Punem intariturile (guseurile) de la flansa de sus + flansele de intrerupere
                    //Cate e pt fiecare mai putin aia de sus
                    if (comboBox_guseu_intreruperi.Text == "12mm")
                    {
                        formula_tabla_12mm = formula_tabla_12mm + arie_guseu_flansa + "*"+(nr_flanse-1)+"*3+";
                    }
                    else if (comboBox_guseu_intreruperi.Text == "10mm")
                    {
                        formula_tabla_10mm = formula_tabla_10mm + arie_guseu_flansa + "*" + (nr_flanse - 1) + "*3+";
                    }
                    else if (comboBox_guseu_intreruperi.Text == "8mm")
                    {
                        formula_tabla_8mm = formula_tabla_8mm + arie_guseu_flansa + "*" + (nr_flanse - 1) + "*3+";
                    }
                }
            }

            MessageBox.Show("Formula tabla de 12mm Length : " + formula_tabla_12mm.Length);
            //Length-1=sterge ultimul caracter dintr-un string
            if (formula_tabla_12mm.Length > 0)
            {
                MessageBox.Show("Formula tabla de 12mm : " + formula_tabla_12mm.Remove(formula_tabla_12mm.Length - 1));
            }

            //Gasesc nr coloanelor pe care se afla UM,greutatea si pretul pt teji
            //Nu prea e in regula artificiul asta sa cauti doar de la coloana C ca sa nu se opreasca la "um" din "nume si sa returneze coloana cu numele dar functioneaza
            bool exista_nr_coloana_UM_tevi = int.TryParse(oe.Gaseste_adresa_text_in_sheet("UM", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "C1:H100", "Coloana"), out nr_coloana_UM_tevi);

            bool exista_nr_coloana_greutate_tevi = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Greutate specifica", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Coloana"), out nr_coloana_greutate_tevi);

            bool exista_nr_coloana_pret_tevi = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Cel mai recent", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Coloana"), out nr_coloana_pret_tevi);

            //Gasesc nr coloanelor pe care se afla UM,greutatea si pretul pt tabla neteda
            //Nu prea e in regula artificiul asta sa cauti doar de la coloana C ca sa nu se opreasca la "um" din "nume si sa returneze coloana cu numele dar functioneaza
            bool exista_nr_coloana_UM_tabla_neteda = int.TryParse(oe.Gaseste_adresa_text_in_sheet("UM", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "C1:H100", "Coloana"), out nr_coloana_UM_tabla_neteda);

            bool exista_nr_coloana_greutate_tabla_neteda = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Greutate specifica", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "A1:H100", "Coloana"), out nr_coloana_greutate_tabla_neteda);

            bool exista_nr_coloana_pret_tabla_neteda = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Cel mai recent", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "A1:H100", "Coloana"), out nr_coloana_pret_tabla_neteda);

            MessageBox.Show("Nr segmente coloana e :" + nr_segmente_coloana);

            if (lungime_fir_coloana_e_numeric == true & lungime_totala_coloana_e_numeric == true)
            {

            }

            teava_stalp_central.Nume = "Pipe " + oe.Gaseste_adresa_text_in_sheet(comboBox_fi_teava.Text, "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Valoare");

            nr_rand_teava_centrala = Int32.Parse(oe.Gaseste_adresa_text_in_sheet(comboBox_fi_teava.Text, "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Rand"));
            nr_rand_teava_sus = Int32.Parse(oe.Gaseste_adresa_text_in_sheet(comboBox_fi_teava_sus.Text, "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Rand"));

            if (exista_nr_coloana_UM_tevi == true)
            {
                teava_stalp_central.Unitate_masura = oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_UM_tevi, nr_rand_teava_centrala);
                teava_sus.Unitate_masura = oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_UM_tevi, nr_rand_teava_sus);
            }

            if (exista_nr_coloana_greutate_tevi == true)
            {
                teava_stalp_central.greutate_specifica = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_greutate_tevi, nr_rand_teava_centrala));
                teava_sus.greutate_specifica = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_greutate_tevi, nr_rand_teava_sus));
            }

            if (exista_nr_coloana_pret_tevi == true)
            {
                teava_stalp_central.Pret = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_pret_tevi, nr_rand_teava_centrala));
                teava_sus.Pret = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", nr_coloana_pret_tevi, nr_rand_teava_sus));
            }

            if ((nr_segmente_coloana > 0) & (lungime_totala_coloana > 0) & (lungime_fir_coloana > 0))
            {
                lungime_segment_coloana = lungime_totala_coloana / nr_segmente_coloana;
                MessageBox.Show("Lungimea segmentului de coloana e" + lungime_segment_coloana);

                if (lungime_fir_coloana < lungime_segment_coloana)
                {
                    MessageBox.Show("Lungimea " + lungime_segment_coloana + " nu poate fi taiata din fir de " + lungime_fir_coloana + "\n Incercati sa folositi alta lungime de fir sau sa mariti numarul de segmente");
                }
                else if ((lungime_fir_coloana >= lungime_segment_coloana) & (teava_stalp_central.greutate_specifica > 0))
                {
                    teava_stalp_central.Formula_excel_cantitate_neta = Returneaza_formula_cantitate_neta_stalp_central(lungime_segment_coloana, nr_intreruperi, teava_stalp_central.Greutate_specifica);
                    //Rotunjim lungime_segment_coloana pana la o valoare divizibila cu lungime_fir_coloana,daca lungime_segment_coloana
                    teava_stalp_central.Formula_excel_cantitate_bruta = Returneaza_formula_cantitate_bruta_stalp_central(lungime_segment_coloana, lungime_fir_coloana, nr_intreruperi);
                    //MessageBox.Show("Cantitatea neta pt teava stalp central:" + Returneaza_formula_cantitate_neta_stalp_central(lungime_segment_coloana, nr_intreruperi, teava_stalp_central.Greutate_specifica));
                    MessageBox.Show("Cantitatea bruta pt teava stalp central:" + Returneaza_formula_cantitate_bruta_stalp_central(lungime_segment_coloana, lungime_fir_coloana, nr_intreruperi));
                }
            }

            if (lungime_teava_sus > 0)
            {
                //poate nu e de 6?
                if (lungime_teava_sus < 6)
                {
                    teava_sus.Formula_excel_cantitate_neta = Returneaza_formula_cantitate_neta_teava_sus(lungime_teava_sus, teava_sus.Greutate_specifica);
                    //nu stiu daca se se taie din 6
                    teava_sus.Formula_excel_cantitate_bruta = Returneaza_formula_cantitate_bruta_teava_sus(lungime_teava_sus, 6);
                    MessageBox.Show("Formula neta calcul teava sus=" + teava_sus.Formula_excel_cantitate_neta);
                    MessageBox.Show("Formula bruta calcul teava sus=" + teava_sus.Formula_excel_cantitate_bruta);
                }
            }

            MessageBox.Show("Teava centrala din: " + teava_stalp_central.Nume + " masurata in " + teava_stalp_central.Unitate_masura + " cu greutatea:" + teava_stalp_central.Greutate_specifica + "si pretul: " + teava_stalp_central.Pret);

            Teava_rotunda placeholder_teava = new Teava_rotunda("test","m",0,0,0,0);
            Surub placeholder_surub = new Surub("test", 0, 0, 0, 0, "test", "test");
            Piulita placeholder_piulita = new Piulita("test", 0, 0, 0, "test", "test");
            Saiba placeholder_saiba = new Saiba("test", 0, 0, 0, "test", "test");
            Saiba_Grower placeholder_saiba_Grower = new Saiba_Grower("test", 0, 0, 0, "test");

            //aici am ramas 06-10-23
            Completeaza_rubrica_stalp_central_spirala(adresa_template_spirala, sheet_template_spirala, teava_stalp_central, placeholder_teava, placeholder_surub, placeholder_piulita, placeholder_saiba, placeholder_saiba_Grower);           

            if ((comboBox_fi_teava.Text == "Ø146x10") & (nr_intreruperi>0))
            {
                teava_bucsa.Nume = "Ø127x8";
                bool exista_nr_rand_teava_bucse = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Ø127x10", adresa_template_spirala, "12.Teava rotunda(Round Pipe)", "A1:H100", "Rand"),out nr_rand_teava_bucse);

                if((exista_nr_rand_teava_bucse==true) &(nr_coloana_UM_tevi>0) & (nr_coloana_greutate_tevi>0) &(nr_coloana_pret_tevi>0))
                {
                    teava_bucsa.Unitate_masura = oe.Returneaza_valoarea_de_la_adresa(adresa_preturi_materiale, "12.Teava rotunda(Round Pipe)", nr_coloana_UM_tevi, nr_rand_teava_bucse);
                    teava_bucsa.Greutate_specifica = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa(adresa_preturi_materiale, "12.Teava rotunda(Round Pipe)", nr_coloana_greutate_tevi, nr_rand_teava_bucse));
                    teava_bucsa.Pret = Convert.ToDouble(oe.Returneaza_valoarea_de_la_adresa(adresa_preturi_materiale, "12.Teava rotunda(Round Pipe)", nr_coloana_pret_tevi, nr_rand_teava_bucse));
                    if((teava_bucsa.Greutate_specifica>0) &(teava_bucsa.Pret>0))
                    {
                        //Bucsele au de obicei o lungime de 0.18m
                        teava_bucsa.Formula_excel_cantitate_neta = "(0.18*" + nr_intreruperi+")*"+teava_bucsa.Greutate_specifica;
                        //Rotunjim 0.18 la 0.2m
                        teava_bucsa.Formula_excel_cantitate_bruta = "("+ Calcule_cantitati_brute.Rotunjeste_lungime_segment(0.18,6)+"*" + nr_intreruperi + ")*" + teava_bucsa.Greutate_specifica;
                        Debug.WriteLine("Formula_excel_cantitate_neta bucsa ="+ teava_bucsa.Formula_excel_cantitate_neta);
                        Debug.WriteLine("Formula_excel_cantitate_bruta bucsa=" + teava_bucsa.Formula_excel_cantitate_bruta);
                        //Debug.WriteLine("Formula_excel_cantitate_neta sus =" + teava_sus.Formula_excel_cantitate_neta);
                    }
                }
            }
            
            //Nu stiu ce bucsa folosesc pt Ø127x8,de la Ø111 in jos,ce gasesc?
            else if ((comboBox_fi_teava.Text == "Ø127x8") & (nr_intreruperi > 0))
            {

            }

            //teava.Calculeaza_suprafata("1+2");

            MessageBox.Show("Avem un nr de " + nr_segmente_coloana + " de segmente de lungime: " + lungime_segment_coloana);

            MessageBox.Show("Rotunjit=" + Calcule_cantitati_brute.Rotunjeste_lungime_segment(6.1, 6));

            MessageBox.Show("Formula pt greutate= " + teava_stalp_central.Calculeaza_suprafata("1+2"));

            String str_rez = Operatiuni_String.Returneaza_string_intre_2_charuri("Pipe Ø 4 2 . 4 x 2 ", 'Ø', 'x');
            //MessageBox.Show("String-ul cu . inlocuit cu , :" + str_rez.Replace('.', ','));
            MessageBox.Show("String-ul fara spatii :" + str_rez.Replace(" ", ""));
            //oe.Schimba_value_celula("S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", 22, 3,"16.00");
            //oe.Insereaza_randuri_worksheet("S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", 23, 2);
            //oe.Copiaza_formatare_celula_in_range("S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala",21,1, 24, 7, 23, 1); ;
            //oe.Sterge_continut_range("S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", 24, 7,23,1);
            //oe.Sterge_randuri_worksheet("S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", 23, 2);
            goto iesi;

            MessageBox.Show("Prima aparitie a cuvantului repos la inceput e pe randul :" + oe.Gaseste_prima_celula_ce_incepe_cu_stringul("Repos", "S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", "A1:H100", "Rand"));

            rand_stalp = Int32.Parse(oe.Gaseste_adresa_text_in_sheet("Pillar", "S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", "A1:H100", "Rand"));
            rand_repos = Int32.Parse(oe.Gaseste_adresa_text_in_sheet("Repos", "S:\\Preturi\\PRETURI MATERIALE\\Calculator spirala\\Calculator_spirala\\Template Spirala.xls", "Spirala", "A1:H100", "Rand"));

            MessageBox.Show("Randul stalpului e:" + rand_stalp);
            MessageBox.Show("Randul reposului e:" + rand_repos);

            Ø_flansa_talpa = Double.Parse(textBox_fi_talpa.Text, CultureInfo.InvariantCulture);

            if (comboBox_gr_talpa.Text == "10mm")
            {
                lista_preturi_tabla_talpa = oe.Gaseste_lista_preturi_din_lista_randuri("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", oe.Gaseste_lista_randuri_caractere_inceput("10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"));
            }
            else if (comboBox_gr_talpa.Text == "8mm")
            {
                lista_preturi_tabla_talpa = oe.Gaseste_lista_preturi_din_lista_randuri("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", oe.Gaseste_lista_randuri_caractere_inceput("8", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"));
            }

            MessageBox.Show("Ø talpa e:" + Ø_flansa_talpa);
            MessageBox.Show("Pretul pt talpa din care e facuta talpa e:" + lista_preturi_tabla_talpa.LastOrDefault().Valoare_RON);


            if ((comboBox_gr_talpa.Text != "10mm") && (comboBox_gr_talpa.Text != "8mm"))
            {
                MessageBox.Show("comboBox_gr_talpa.Text=" + comboBox_gr_talpa.Text);
                MessageBox.Show("Grosime flansa talpa necunoscuta");
            }

            // | e OR logic in C#
            if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_fir.Text) == false)
            {
                MessageBox.Show("Lungime fir din care se taie coloana stalpului central necunoscuta");
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_totala.Text) == false)
            {
                MessageBox.Show("Lungime totala coloana centrala necunoscuta");
            }
            else if ((comboBox_fi_teava.Text != "Ø146x10") && (comboBox_fi_teava.Text != "Ø127x8"))
            {
                MessageBox.Show("Ø stalp central necunoscut");
            }
            else if ((comboBox_gr_talpa.Text != "10mm") && (comboBox_gr_talpa.Text != "8mm"))
            {
                MessageBox.Show("Grosime flansa talpa necunoscuta");
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_fir.Text) == false)
            {
                MessageBox.Show("Diametru flansa talpa necunoscut");
            }
            else if ((comboBox_fi_teava_sus.Text != "Ø48x3") && (comboBox_fi_teava_sus.Text != "Ø60.3x3"))
            {
                MessageBox.Show("Diametru teava sus necunoscut");
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_L_teava_sus.Text) == false)
            {
                MessageBox.Show("Lungime teava sus necunoscuta");
            }
            //Daca rubrica nr intreruperi are un nr > 0 -> avem intreruperi,deci trebuie sa le verificam 
            else if (Verificari.Verifica_daca_un_string_e_integer(textBox_nr_Intreruperi.Text) == true)
            {
                if (Convert.ToInt16(textBox_nr_Intreruperi.Text) > 0)
                {

                    if (Verificari.Verifica_daca_un_string_e_double(textBox_fi_intreruperi.Text) == false)
                    {
                        MessageBox.Show("Ø flansa intrerupere necunoscuta");
                    }
                    else if ((comboBox_gr_intreruperi.Text != "10mm") && (comboBox_gr_intreruperi.Text != "8mm"))
                    {
                        MessageBox.Show("Grosime flansa intrerupere necunoscuta");
                    }
                    else if ((comboBox_gr_intreruperi.Text != "10mm") && (comboBox_gr_intreruperi.Text != "8mm"))
                    {
                        MessageBox.Show("Diametru teava sus necunoscut");
                    }
                    else if (Convert.ToInt16(textBox_fi_intreruperi.Text) >= Convert.ToInt16(textBox_fi_talpa.Text))
                    {
                        MessageBox.Show("Ø flansa intrerupere > Ø flansa baza");
                    }
                    //Daca nu se indeplinesc conditiile de mai sus-> formularul a fost completat bine deci putem calcula
                    else
                    {
                        nr_intreruperi = Convert.ToInt16(textBox_nr_Intreruperi.Text);

                        lungime_totala_coloana = Convert.ToDouble(textBox_lungime_totala.Text);
                        lungime_fir_coloana = Convert.ToDouble(textBox_lungime_fir);
                        lungime_segment_coloana = lungime_totala_coloana / nr_intreruperi;
                        if (lungime_segment_coloana > lungime_totala_coloana)
                        {
                            MessageBox.Show("Lungimea unui segment de coloana: " + lungime_segment_coloana + " depaseste lungimea firului de " + lungime_fir_coloana + "din care se taie ");
                        }
                        else
                        {
                            Stalpul_se_poate_calcula = true;
                        }
                    }
                }
            }
            else
            {
                lungime_totala_coloana = Convert.ToDouble(textBox_lungime_totala.Text);
                lungime_segment_coloana = lungime_totala_coloana;
                lungime_fir_coloana = Convert.ToDouble(textBox_lungime_fir);

                if (lungime_segment_coloana > lungime_fir_coloana)
                {
                    MessageBox.Show("Lungimea unui segment de coloana: " + lungime_segment_coloana + " depaseste lungimea firului de " + lungime_fir_coloana + "din care se taie ");
                }
                else
                {
                    //Daca sunt indeplinite toate conditiile inseamna ca form-ul Stalp_central a fost completat bine deci putem memora datele din form in variabile pt a putea fi folosite
                    Stalpul_se_poate_calcula = true;
                }

            }

            //asta e doar pt verificari
            if (Verificari.Verifica_daca_un_string_e_integer(textBox_nr_Intreruperi.Text) == true)
            {
                if (Convert.ToInt16(textBox_nr_Intreruperi.Text) > 0)
                {
                    nr_intreruperi = Convert.ToInt16(textBox_nr_Intreruperi.Text);

                    if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_totala.Text) == true)
                    {

                        lungime_segment_coloana = Convert.ToDouble(textBox_lungime_totala.Text) / nr_intreruperi;
                        MessageBox.Show("Lungime segment coloana=" + lungime_segment_coloana);
                        //goto iesi;
                    }
                }
            }

            MessageBox.Show("Pt 10 am gasit:" + oe.Gaseste_lista_randuri_caractere_inceput("10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"));


            //MessageBox.Show(os.Elimina_un_substring_din_string(oe.Gaseste_adresa_text_in_sheet("Ø146x10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Adresa"), "$"));

            coloana_test = Int32.Parse(oe.Gaseste_adresa_text_in_sheet("cel mai recent", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Coloana"));
            rand_test = Int32.Parse(oe.Gaseste_adresa_text_in_sheet("Ø146x10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", "A1:H100", "Rand"));

        //asta e doar pt testate
        iesi:
            this.Close();

            //MessageBox.Show("Primele 2 caractere din 10x1500x6000/S355=" + os.Returneaza_caractere_inceput_string("10x1500x6000/S355",2));

            //MessageBox.Show("Rand test=" + rand_test);

            //MessageBox.Show("Textbox lungime fir e=" + textBox_lungime_fir.Text);

            //MessageBox.Show("Pretul Ø146x10:" + oe.Returneaza_valoarea_de_la_adresa("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "12.Teava rotunda(Round Pipe)", coloana_test, rand_test));
        }

        //Butonul "Anuleaza"-> Va inchide form-ul care e in prezent deschis
        private void buton_anuleaza_Click(object sender, EventArgs e)
        {
            //Verificari v = new Verificari();
            Operatiuni_Excel oe = new Operatiuni_Excel();
            Operatiuni_Lista ol = new Operatiuni_Lista();

            string adresa_Surub_piulita = "";

            Surub test_M8 = new Surub("Surub cap inecat", 0, 0, 8, 30, "inecat", "A2");
            Piulita test_piulita_M8 = new Piulita("Piulite hexagonale normale", 0, 0, 8, "hexagonala", "A2");
            Saiba test_saiba_M8 = new Saiba("Saiba plata pt metale", 0, 0, 8, "plata", "A2");
            Saiba_Grower test_saiba_Grower_M8 = new Saiba_Grower("Saiba plata pt metale", 0, 0, 8, "A2");

            MessageBox.Show("Greutatea piulitei de M8 gasite e " + test_saiba_Grower_M8.Extrage_greutate_din_tabel_suruburi("Piulita hexagonala M8", "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"));
            MessageBox.Show("Pretul piulitei de M8 finisaj inox gasite e " + test_saiba_Grower_M8.Extrage_pret_din_tabel_suruburi("Piulita hexagonala M8", "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"));

            //MessageBox.Show("Greutatea surbului gasit e:" + test_M8.Extrage_greutate_din_tabel_suruburi(test_M8.Tip_cap, "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"));
            //MessageBox.Show("Pretul surbului gasit e:" + test_M8.Extrage_pret_din_tabel_suruburi(test_M8.Tip_cap, "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx"));

            //bool exista_nr_coloana_surub_cap_hexagonal = int.TryParse(oe.Gaseste_adresa_text_in_sheet("Surub cap hexagonal+piulita", "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsm", "M16", "B1:H100", "Adresa"), out adresa_Surub_piulita);

            int col_test = int.Parse(oe.Gaseste_adresa_text_in_sheet("Surub cap hexagonal+piulita", "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx", "M16", "B1:H100", "Rand"));
            MessageBox.Show("Coloana Surub cap hexagonal+piulita e:" + oe.Gaseste_adresa_text_in_sheet("60", "S:\\Preturi\\PRETURI MATERIALE\\SURUBURI.xlsx", "M16", "B" + col_test + ":H100", "Adresa"));

            List<Segment> lista_segmente = ol.Creaza_lista_de_segmente_noua();
            Segment seg1 = new Segment(1.2, 3);
            Segment seg2 = new Segment(1.5, 1);
            Segment seg3 = new Segment(2, 2);

            lista_segmente = ol.Adauga_un_segment_in_lista_de_segmente(lista_segmente, seg1);
            lista_segmente = ol.Adauga_un_segment_in_lista_de_segmente(lista_segmente, seg2);
            lista_segmente = ol.Adauga_un_segment_in_lista_de_segmente(lista_segmente, seg3);

            MessageBox.Show("Formula obtinuta: " + ol.Transforma_lista_segmente_in_string(lista_segmente));

            //List<Pret> lista_preturi = new List<Pret>();

            //if (Verificari.Verifica_daca_o_lista_de_preturi_nu_e_goala(oe.Gaseste_lista_preturi_din_lista_randuri("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", oe.Gaseste_lista_randuri_caractere_inceput("10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"))))
            //{
            //    lista_preturi = oe.Gaseste_lista_preturi_din_lista_randuri("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", oe.Gaseste_lista_randuri_caractere_inceput("10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"));

            //    //Pt a sorta descendent
            //    var lista_noua = lista_preturi.OrderByDescending(x => x.Data_primire).ToList();
            //    //lista_preturi.Sort((a, b) => b.CompareTo(a));

            //    for (var i = 0; i < lista_noua.Count; i++)
            //    {
            //        Debug.WriteLine("Data nr " +i+ " "+ lista_noua[i].Data_primire);
            //        Debug.WriteLine("Pretul nr " + i + " [RON] " + lista_noua[i].Valoare_RON);
            //    }
            //}

            //oe.Gaseste_lista_preturi_din_lista_randuri("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", oe.Gaseste_lista_randuri_caractere_inceput("10", "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40"));

            //TODO:Sa fac cumva sa separ tabla S355 de restul?
            //MessageBox.Show("Cel mai nou pret la tabla de 10 e :" + oe.Returneaza_cel_mai_recent_pret("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm", "8.Tabla neteda", "B5:B40", "10"));

            MessageBox.Show("Ø12 de 1m " + Calcule_greutati.Calculeaza_greutate_bara_rotunda(12, 1, 7850));
            MessageBox.Show("Bara patrata de 12 de 1m " + Calcule_greutati.Calculeaza_greutate_bara_dreptunghiulara(12, 12, 1, 7850));
            //MessageBox.Show("Preturi material e deschis:" + v.Verifica_daca_workbook_e_deschis("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm"));
            //MessageBox.Show("Workbook-ul gasit:" + oe.Deschide_workbook_dupa_path("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm"));
            //if(oe.Deschide_workbook_dupa_path("S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm") is not null)
            //{
            //    MessageBox.Show("I should be seeing this");
            //}
            this.Close();
        }

        //Returneaza un string ce poate fi folosit ca o formula in Excel in functie de greutatea/m a stalpului central si de nr de segmente:
        private string Returneaza_formula_cantitate_neta_stalp_central(double lungime_segment, int nr_intreruperi, double greutate_specifica_teava)
        {
            int nr_segmente_stalp_central = 0;

            if ((lungime_segment > 0) & (greutate_specifica_teava > 0))
            {
                //Daca avem intreruperi trebuie sa returnam o formula pt Excel de forma : "=(lungime_segment*nr_segmente)*greutate_specifica_teava"
                if (nr_intreruperi > 0)
                {
                    nr_segmente_stalp_central = nr_intreruperi + 1;
                    return "=(" + lungime_segment.ToString() + "*" + nr_segmente_stalp_central + ")*" + greutate_specifica_teava.ToString();
                }
                //Daca nu avem intreruperi trebuie sa returnam o formula pt Excel de forma : "=lungime_segment*greutate_specifica_teava"
                else if (nr_intreruperi == 0)
                {
                    return "=" + lungime_segment.ToString() + "*" + greutate_specifica_teava.ToString();
                }
            }
            return "";
        }

        private string Returneaza_formula_cantitate_bruta_stalp_central(double lungime_segment, double lungime_fir, int nr_intreruperi)
        {
            int nr_segmente_stalp_central = 0;

            if ((lungime_fir > 0) & (lungime_segment > 0))
            {
                //Rotunjim lungime_segment pana la o valoare divizibila cu lungime_fir
                lungime_segment = Calcule_cantitati_brute.Rotunjeste_lungime_segment(lungime_segment, lungime_fir);
                //Daca avem intreruperi trebuie sa returnam o formula pt Excel de forma : "=(lungime_segment*nr_segmente)*greutate_specifica_teava"
                if (nr_intreruperi > 0)
                {
                    nr_segmente_stalp_central = nr_intreruperi + 1;
                    return "=(" + lungime_segment.ToString() + "*" + nr_segmente_stalp_central + ")";
                }
                //Daca nu avem intreruperi trebuie sa returnam o formula pt Excel de forma : "=lungime_segment*greutate_specifica_teava"
                else if (nr_intreruperi == 0)
                {
                    return "=" + lungime_segment.ToString();
                }
            }
            return "";
        }

        //Returneaza un string ce poate fi folosit ca o formula in Excel in functie de greutatea/m a tejii de sus
        private string Returneaza_formula_cantitate_neta_teava_sus(double lungime_segment, double greutate_specifica_teava)
        {
            if ((lungime_segment > 0) & (greutate_specifica_teava > 0))
            {
                return "=" + lungime_segment.ToString() + "*" + greutate_specifica_teava.ToString();
            }

            return "";
        }

        private string Returneaza_formula_cantitate_bruta_teava_sus(double lungime_segment, double lungime_fir_din_care_se_taie)
        {
            if ((lungime_segment > 0) & (lungime_fir_din_care_se_taie > 0))
            {
                if (lungime_segment <= lungime_fir_din_care_se_taie)
                {
                    lungime_segment = Calcule_cantitati_brute.Rotunjeste_lungime_segment(lungime_segment, lungime_fir_din_care_se_taie);
                    return "=" + lungime_segment.ToString();
                }
            }

            return "";
        }

        //Retuneaza aria guseului de la talpa stalpului (in m2)
        //diametru_stalp_central si Ø_placa_talpa se dau in mm
        private double Returneaza_arie_guseu_talpa(string nume_teava_stalp_central, double Ø_placa_talpa)
        {
            double diametru_obtinut = 0;

            //Daca e<2 nu putem string-ul dintre caracterle 
            if (nume_teava_stalp_central.Length > 2)
            {
                nume_teava_stalp_central = Operatiuni_String.Returneaza_string_intre_2_charuri(nume_teava_stalp_central, 'Ø', 'x');
            }

            bool e_diametru_valabil = double.TryParse(nume_teava_stalp_central, out diametru_obtinut);

            if ((diametru_obtinut > 0) & (Ø_placa_talpa > 0))
            {
                if (Ø_placa_talpa <= diametru_obtinut + 63.3 * 2)
                {
                    return 0;
                }
                else if (Ø_placa_talpa >= diametru_obtinut + 63.3 * 2)
                {
                    //trebuie decimal mai intai altfel 20/1000 da direct zero
                    decimal inaltime_dreptunghi_jos = 20;
                    decimal inaltime_dreptunghi_sus = 40;
                    decimal cateta_triunghi_sus = 63.3M;

                    //22 pt ca asa am luat din desen
                    Patrulater dreptunghi_jos = new Patrulater((Ø_placa_talpa - diametru_obtinut) / 2 / 1000, decimal.ToDouble(inaltime_dreptunghi_jos) / 1000);
                    //40 pt ca asa am luat din desen
                    Patrulater dreptunghi_sus = new Patrulater((Ø_placa_talpa - diametru_obtinut - 63.3 * 2) / 2 / 1000, decimal.ToDouble(inaltime_dreptunghi_sus) / 1000);
                    //40 pt ca asa am luat din desen
                    Triunghi_dreptunghic triunghi_sus = new Triunghi_dreptunghic(decimal.ToDouble(cateta_triunghi_sus) / 1000, decimal.ToDouble(inaltime_dreptunghi_sus) / 1000);

                    Debug.WriteLine("Arie dreptunghi jos :" + dreptunghi_jos.Calculeaza_arie());
                    Debug.WriteLine("Arie dreptunghi sus :" + dreptunghi_sus.Calculeaza_arie());
                    Debug.WriteLine("Arie triunghi sus :" + triunghi_sus.Calculeaza_arie());

                    return dreptunghi_jos.Calculeaza_arie() + dreptunghi_sus.Calculeaza_arie() + triunghi_sus.Calculeaza_arie();
                }
            }
            return 0;
        }

        //Retuneaza aria guseului flansei de sus si celei pt intreruperi (in m2)
        //diametru_stalp_central si Ø_flansa se dau in mm
        private double Returneaza_arie_guseu_flansa(string nume_teava_stalp_central, double Ø_flansa)
        {
            double diametru_obtinut = 0;

            //Daca e<2 nu putem string-ul dintre caracterle 
            if (nume_teava_stalp_central.Length > 2)
            {
                nume_teava_stalp_central = Operatiuni_String.Returneaza_string_intre_2_charuri(nume_teava_stalp_central, 'Ø', 'x');
            }

            bool e_diametru_valabil = double.TryParse(nume_teava_stalp_central, out diametru_obtinut);

            if ((diametru_obtinut > 0) & (Ø_flansa > 0))
            {
                if (Ø_flansa <= diametru_obtinut)
                {
                    return 0;
                }
                else if (Ø_flansa > diametru_obtinut)
                {
                    //trebuie decimal mai intai altfel 20/1000 da direct zero
                    decimal inaltime_flansa = 67;

                    Triunghi_dreptunghic triunghi_flansa = new Triunghi_dreptunghic((Ø_flansa - diametru_obtinut) / 2 / 1000, decimal.ToDouble(inaltime_flansa) / 1000);

                    Debug.WriteLine("Arie flansa :" + triunghi_flansa.Calculeaza_arie());

                    return triunghi_flansa.Calculeaza_arie();
                }
            }

            return 0;
        }

        //Verifica cate grosimi de talpa exista
        //1-daca avem 12mm peste tot (guseuri,flansa talpa,flansa sus)
        //2-daca avem 2 grosimi (guseuri,flansa talpa,flansa sus)
        //3-daca avem 3 grosimi (avem alte grosimi in afara de 12,10,8mm ?) (guseuri,flansa talpa,flansa sus)
        private int Cate_grosimi_de_tabla_avem()
        {
            string[] grosimi_tabla = { "12mm", "10mm", "8mm" };

            //Vedem ce grosimi avem in ComboBox
            //12,10 si 8mm (se mai folosesc si altele?)
            int nr_grosimi_nefolosite = grosimi_tabla.Length;

            //Grosimea selectata in form va fi eliminata din array-ul grosimi_tabla
            grosimi_tabla = Operatiuni_Array.Sterge_string_din_array_string(grosimi_tabla, comboBox_gr_intreruperi.Text);
            grosimi_tabla = Operatiuni_Array.Sterge_string_din_array_string(grosimi_tabla, comboBox_gr_talpa.Text);
            grosimi_tabla = Operatiuni_Array.Sterge_string_din_array_string(grosimi_tabla, comboBox_guseu_intreruperi.Text);
            grosimi_tabla = Operatiuni_Array.Sterge_string_din_array_string(grosimi_tabla, comboBox_guseu_talpa.Text);

            Debug.WriteLine("Lungime sir="+ grosimi_tabla.Length);

            return (nr_grosimi_nefolosite- grosimi_tabla.Length);
        }

        //Aici am ramas 05-10-23
        //Numarul de materiale (teji,table,suruburi,etc) necesare pt a fabrica stalpul
        //Incepem cu un minim de 8 materiale:
        // - teava pt stalpul central
        // - teava sus
        // - capac pt teava sus (tabla 3mm)
        // - tabla pt flanse+guseuri
        //-surub pt prindere flansa cu teava sus de flansa cu stalpul central
        //Trebuie lamurit!
        //-surub pt prindere flansa cu teava sus de flansa cu stalpul central (partea cu tabla de 20mm)
        //-piulita pt prindere flansa cu teava sus de flansa cu stalpul central
        //-saiba plata pt prindere flansa cu teava sus de flansa cu stalpul central
        //-saiba Grower pt prindere flansa cu teava sus de flansa cu stalpul central
        private int Returneaza_nr_de_materiale_folosit()
        {
            int numar_materiale = 8;
            //Verificam daca nr_intreruperi a fost completat cu un nr intreg
            bool nr_intreruperi_e_numeric = false;
            int nr_intreruperi = 0;

            nr_intreruperi_e_numeric = int.TryParse(textBox_nr_Intreruperi.Text, out nr_intreruperi);

            if(nr_intreruperi_e_numeric==true)
            {
                nr_intreruperi = int.Parse(textBox_nr_Intreruperi.Text);
                //  Pt cazul in care avem stalp din Ø127x8,de obicei nu se pun bucsi  
                if ((nr_intreruperi>0) & comboBox_fi_teava.Text== "Ø146x10")
                {
                    numar_materiale = numar_materiale + 1;
                }
            }

            //Grosimi de tabla
            if(Cate_grosimi_de_tabla_avem()!=1)
            {
                numar_materiale = numar_materiale + (Cate_grosimi_de_tabla_avem() - 1);
            }

            return numar_materiale;
        }

        //Aici am ramas 05-10-2023
        //Bool care verifica daca form-ul "Stalp_central" e completat bine adica:
        //Stalpul central are o lungime totala si o lungime de fir
        //lungime de fir>lungime totala pt stalp fara intreruperi SAU
        //lungime de fir>lungime totala/nr_intreruperi pt stalp cu intreruperi
        //Stalpul de sus are o lungime
        //Ø_flansa are o valoare
        //Ø_flansa_talpa are o valoare
        //int nr_intr (nr intreruperi) = 0 face parametrul optional
        //Daca nr de intreruperi e mai mare ca 0
        private bool Verifica_daca_formul_e_completat_bine(int nr_intr,double lung_tot_col, double lung_fir_col)
        {
            //Aici vom memora lungimea segmentului de coloana care poate fi: =lung_tot_col sau lung_tot_col/nr_intr in cazul in care spirala e impartita in segmente
            double lung_seg_col = 0;

            //if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_fir.Text) == false)
            if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_fir.Text) == false)
            {
                MessageBox.Show("Lungime fir din care se taie coloana stalpului central necunoscuta");
                return false;
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_totala.Text) == false)
            {
                MessageBox.Show("Lungime totala coloana centrala necunoscuta");
                return false;
            }
            else if ((comboBox_fi_teava.Text != "Ø146x10") && (comboBox_fi_teava.Text != "Ø127x8"))
            {
                MessageBox.Show("Ø stalp central necunoscut");
                return false;
            }
            else if ((comboBox_gr_talpa.Text != "12mm") && (comboBox_gr_talpa.Text != "10mm") && (comboBox_gr_talpa.Text != "8mm"))
            {
                MessageBox.Show("Grosime flansa talpa necunoscuta");
                return false;
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_lungime_fir.Text) == false)
            {
                MessageBox.Show("Diametru flansa talpa necunoscut");
                return false;
            }
            else if ((comboBox_fi_teava_sus.Text != "Ø48.3x3") && (comboBox_fi_teava_sus.Text != "Ø48.3x4") && (comboBox_fi_teava_sus.Text != "Ø60.3x3"))
            {
                MessageBox.Show("Diametru teava sus necunoscut");
                return false;
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_L_teava_sus.Text) == false)
            {
                MessageBox.Show("Lungime teava sus necunoscuta");
                return false;
            }
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_L_teava_sus.Text) == false)
            {
                MessageBox.Show("Lungime teava sus necunoscuta");
                return false;
            }
            //Verificam ca firul din care se taie teava de sus sa fie <6m
            else if (Verificari.Verifica_daca_un_string_e_double(textBox_L_teava_sus.Text) == true)
            {
                double lungime_teava_sus = Double.Parse(textBox_L_teava_sus.Text);
                //de obicei firul se aduce la 6m
                //l-am pus asa ca nu mai adaug o rubrica "Lungime fir [m]"
                if (lungime_teava_sus > 6)
                {
                    return false;
                }
                //goto Continua;
            }

        //Continua:
            if (Verificari.Verifica_daca_un_string_e_double(textBox_fi_intreruperi.Text) == false)
            {
                MessageBox.Show("Ø flansa sus/intermediara necunoscuta");
                return false;
            }
            else if ((comboBox_gr_intreruperi.Text != "12mm") && (comboBox_gr_intreruperi.Text != "10mm") && (comboBox_gr_intreruperi.Text != "8mm"))
            {
                MessageBox.Show("Grosime flansa intrerupere necunoscuta");
                return false;
            }

            //Daca rubrica nr intreruperi are un nr > 0 -> avem intreruperi,deci trebuie sa le verificam 
            else if (Verificari.Verifica_daca_un_string_e_integer(textBox_nr_Intreruperi.Text) == true)
            {
                if (Convert.ToInt16(textBox_nr_Intreruperi.Text) > 0)
                {
                    if ((comboBox_gr_intreruperi.Text != "12mm") && (comboBox_gr_intreruperi.Text != "10mm") && (comboBox_gr_intreruperi.Text != "8mm"))
                    {
                        MessageBox.Show("Diametru teava sus necunoscut");
                        return false;
                    }
                    else if (Convert.ToInt16(textBox_fi_intreruperi.Text) >= Convert.ToInt16(textBox_fi_talpa.Text))
                    {
                        MessageBox.Show("Ø flansa intrerupere > Ø flansa baza");
                        return false;
                    }
                    //Daca nu se indeplinesc conditiile de mai sus-> formularul a fost completat bine deci putem calcula
                    else
                    {
                        //nr_intr = Convert.ToInt16(textBox_nr_Intreruperi.Text);

                        //lungime_totala_coloana = Convert.ToDouble(textBox_lungime_totala.Text);
                        //lungime_fir_coloana = Convert.ToDouble(textBox_lungime_fir);
                        //lungime_segment_coloana = lungime_totala_coloana / nr_intreruperi;
                        if (lung_seg_col > lung_tot_col)
                        {
                            MessageBox.Show("Lungimea unui segment de coloana: " + lung_seg_col + " depaseste lungimea firului de " + lung_fir_col + "din care se taie ");
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
            }
            else
            {
                //lungime_totala_coloana = Convert.ToDouble(textBox_lungime_totala.Text);
                //lungime_segment_coloana = lungime_totala_coloana;
                //lungime_fir_coloana = Convert.ToDouble(textBox_lungime_fir);

                if (lung_seg_col > lung_fir_col)
                {
                    MessageBox.Show("Lungimea unui segment de coloana: " + lung_seg_col + " depaseste lungimea firului de " + lung_fir_col + "din care se taie ");
                    return false;
                }
                else
                {
                    //Daca sunt indeplinite toate conditiile inseamna ca form-ul Stalp_central a fost completat bine deci putem memora datele din form in variabile pt a putea fi folosite
                    return true;
                }
            }
            return false;
        }

        //aici am ramas 05-10-23
        //Completez celulele Excel intre rubricile "Pillar" si "Repos" cu consumul de materiale necesar
        private void Completeaza_rubrica_stalp_central_spirala(string adresa_template_spirala,string sheet_template_spirala,Teava_rotunda tv_stalp_central,Teava_rotunda tv_teava_sus, Surub Surub_prindere_flanse,Piulita Piulita_prindere, Saiba Saiba_prindere,Saiba_Grower Saiba_Grower_prindere, Surub Surub_prindere_repos=null, Teava_rotunda tv_bucsa=null,string formula_12mm="", string formula_10mm="", string formula_8mm="")
        {
            //Pregatim tabelul pt a insera celulele:
            Operatiuni_Excel oe = new Operatiuni_Excel();
            //gasim randul in care scrie "Pillar" (rand_inceput_pillar) si randul in care scrie "Repos" (rand_sfarsit_pillar)
            int rand_inceput_pillar = Int32.Parse(oe.Gaseste_informatii_text_in_sector_sheet("Pillar", "Raw material", adresa_template_spirala, sheet_template_spirala, "A1:G100", "Rand"));
            int rand_sfarsit_pillar = Int32.Parse(oe.Gaseste_informatii_text_in_sector_sheet("Repos", "Raw material", adresa_template_spirala, sheet_template_spirala, "A1:G100", "Rand"));

            //Stergem prima celula sub Denumirea "Pillar"

            //am scazut -2 pt ca:
            //-1 pt ca rand_sfarsit_pillar-rand_inceput_pillar da nr celule intre "Pillar" si "Repos"+1
            //-1 pt a pastra o celula si a nu pierde formatarea
            int nr_randuri_de_sters = rand_sfarsit_pillar-rand_inceput_pillar - 2;
            //Stergem tot continutul din primul rand (ne intereseaza doar sa pastram formatarea)
            oe.Sterge_continut_range(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 1, 1, rand_inceput_pillar + 1, 7);
            //Stergem tot in afara de primul rand golit de continut
            oe.Sterge_randuri_worksheet(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, nr_randuri_de_sters);
            //Inseram nr de celule necesar
            oe.Insereaza_randuri_worksheet(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, Returneaza_nr_de_materiale_folosit());
            //Adaugam chenare in randurile nou introduse
            oe.Adauga_chenare_pt_celula_in_range(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 1, rand_inceput_pillar + 2 + Returneaza_nr_de_materiale_folosit(), 7);

            //Aici am ramas 06-10-2023
            //Daca stalp_central !=null
            if (tv_stalp_central != null)
            {
                if (string.IsNullOrWhiteSpace(tv_stalp_central.Nume) == false)
                {
                    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 1, tv_stalp_central.Nume);
                }

                if (string.IsNullOrWhiteSpace(tv_stalp_central.Unitate_masura) == false)
                {
                    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 2,tv_stalp_central.Unitate_masura);
                }

                if (string.IsNullOrWhiteSpace(tv_stalp_central.Formula_excel_cantitate_bruta) == false)
                {
                    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 3, tv_stalp_central.Formula_excel_cantitate_bruta);
                }

                if (tv_stalp_central.Pret>0)
                {
                    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 3, tv_stalp_central.Formula_excel_cantitate_bruta);
                }

                if (string.IsNullOrWhiteSpace(tv_stalp_central.Formula_excel_cantitate_neta) == false)
                {
                    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 6, tv_stalp_central.Formula_excel_cantitate_neta);
                }

                //Trebuie calculator formula suprafata
                //if (string.IsNullOrWhiteSpace(tv_stalp_central.) == false)
                //{
                //    oe.Schimba_value_celula(adresa_template_spirala, sheet_template_spirala, rand_inceput_pillar + 2, 6, tv_stalp_central.Formula_excel_cantitate_neta);
                //}
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }
    }
}
