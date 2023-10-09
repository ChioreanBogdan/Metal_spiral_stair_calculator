using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
//using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Calculator_spirala.Obiecte;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace Calculator_spirala.Modules
{
    internal class Operatiuni_Excel
    {
        public void Deschide_preturi_materiale()
        {
            string de_deschis = @"S:\Preturi\PRETURI MATERIALE\PRETURI MATERIALE.xlsm";
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbooks books = excelApp.Workbooks;

            Excel.Workbook sheet = books.Open(de_deschis);
        }

        public void Deschide_preturi(string sheet_de_deschis)
        {
            Microsoft.Office.Interop.Excel.Worksheet Sheet_gasit;

            Excel.Application excelApp = new Excel.Application();

            // if you want to make excel visible to user, set this property to true, false by default
            excelApp.Visible = true;

            // open an existing workbook
            string workbookPath = "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm";
            Excel.Workbook Workbook_preturi = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);

            //Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkbook.Sheets[sheet_de_deschis];

            //excelWorkbook.Activate();
            
            excelApp.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

            Sheet_gasit = excelApp.ActiveWorkbook.Sheets[sheet_de_deschis];

            //MessageBox.Show(excelApp.ActiveWorkbook.Sheets[sheet_de_deschis].Range("B3").Value2);
        }

        //Transforma un nr de coloana in cifra corespunzatoare:
        //Ex: Col cu nr 1 ->A 2->B,etc
        public string Returneaza_litera_coloana_dupa_nr(int nr_coloana)
        {
            string Litera_coloana = "";

            while (nr_coloana > 0)
            {
                int modulo = (nr_coloana - 1) % 26;
                Litera_coloana = Convert.ToChar('A' + modulo) + Litera_coloana;
                nr_coloana = (nr_coloana - modulo) / 26;
            }

            return Litera_coloana;
        }

        //Memoreaza intr-o lista toate nr de randuri in care apare un text ce incepe cu un anumit caracter/sir de caractere (caractere_cautate)
        //De ex: Returneaza_caractere_inceput_string("10x1500x6000/S355",2) retrneaza "10" -> se va adauga in lista_randuri nr randului/randurilor unde a fost gasita "10x1500x6000/S355"
        public List<int> Gaseste_lista_randuri_caractere_inceput(string caractere_cautate,string workbook_de_deschis, string sheet_de_deschis, string array_de_cautare)
        {
            string File_name = workbook_de_deschis;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            try
            {
                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                Excel.Range currentFind = null;
                Excel.Range firstFind = null;
                int nr_gasit = 0;
                //Aici vom memora,sub forma de int-uri,toate randurile ce contin la inceput string-ul "caractere_cautate"
                List<int> lista_randuri = new List<int>();

                //Operatiuni_String os = new Operatiuni_String();

                // You should specify all these parameters every time you call this method,
                // since they can be overridden in the user interface.
                currentFind = oSheet.Range[array_de_cautare].Find(caractere_cautate, missing,
                    Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                    missing, missing);

                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                          == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }


                    if (Operatiuni_String.Returneaza_caractere_inceput_string(currentFind.Value2, caractere_cautate.Length) == caractere_cautate)
                    {
                        nr_gasit++;
                        lista_randuri.Add(currentFind.Row);
                        //MessageBox.Show("Current find is:" + currentFind.Value2);
                    }

                   currentFind = oSheet.Range[array_de_cautare].FindNext(currentFind);
                }

                //Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange("10", array_de_cautare, oSheet);
                if (nr_gasit != 0)
                {
                    //MessageBox.Show("Nr de aparitii a lui 10:"+ nr_gasit);

                    int i = 0;

                    if (lista_randuri.Count>0)
                    {
                        //foreach (int ran in lista_randuri)
                        //{
                        //    MessageBox.Show("coloana[" + i + "] gasita=" + ran);
                        //    i++;
                        //}
                        oWB.Close(false, missing, missing);

                        oSheet = null;
                        oWB = null;
                        oXL.Quit();

                        return lista_randuri;
                    }

                }
                else
                {
                    oWB.Close(false, missing, missing);

                    oSheet = null;
                    oXL.Quit();
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            //int[] Randuri_gasite;
            //Randuri_gasite = new int[3];
            return null;
        }

        //Gaseste prima celula ce incepe cu string-ul string_cautat si,returneaza sub forma de string adresa,randul,coloana sau continutul acestia,in functie de ce se cere in string-ul "ce_caut"
        public string Gaseste_prima_celula_ce_incepe_cu_stringul(string string_cautat, string workbook_de_deschis, string sheet_de_deschis, string array_de_cautare, string ce_caut)
        {
            //Aici memoram,sub forma de string, adresa,randul,coloana sau continutul primei celule ce incepe cu caracterele string_cautat
            string valoare_gasita = "";

            string File_name = workbook_de_deschis;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            try
            {
                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                //Excel.Range rng = (Excel.Range)oSheet.Cells[21, 1];

                //MessageBox.Show("oSheet.Cells[21, 1]=" + rng.Row);
                //MessageBox.Show("oSheet.Cells[21, 1]=" + rng.Value2);

                Excel.Range currentFind = null;
                Excel.Range firstFind = null;

                Excel.Range Range_gasit = null;

                // You should specify all these parameters every time you call this method,
                // since they can be overridden in the user interface.
                currentFind = oSheet.Range[array_de_cautare].Find(string_cautat, missing,
                    Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                    missing, missing);

                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                          == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    if (Operatiuni_String.Returneaza_caractere_inceput_string(currentFind.Value2, string_cautat.Length) == string_cautat)
                    {
                        Range_gasit = currentFind;
                        //ne oprim dupa primul range gasit
                        goto a_fost_gasit_un_range;
                        //MessageBox.Show("Current find is:" + currentFind.Value2);
                    }

                    currentFind = oSheet.Range[array_de_cautare].FindNext(currentFind);
                }

                a_fost_gasit_un_range:

                //Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange("10", array_de_cautare, oSheet);
                if (Range_gasit != null)
                {
                    switch (ce_caut)
                    {
                        case "Adresa":
                            valoare_gasita = Range_gasit.Address;
                            break;
                        case "Coloana":
                            valoare_gasita = Range_gasit.Column.ToString();
                            break;
                        case "Rand":
                            valoare_gasita = Range_gasit.Row.ToString();
                            break;
                        case "Valoare":
                            valoare_gasita = Range_gasit.Value2.ToString();
                            break;
                    }

                    oWB.Close(false, missing, missing);

                        oSheet = null;
                        oWB = null;
                        oXL.Quit();

                        return valoare_gasita;
                }

                else
                {
                    oWB.Close(false, missing, missing);

                    oSheet = null;
                    oXL.Quit();
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return null;
        }

        //Cauta un text in sheet si,daca il gaseste,returneaza adresa celulei in care se gaseste textul
        //IMPORTANT: daca sunt mai multe string-uri identice acesta va returna adresa primului string gasit
        //array_de_cautare: in ce array de celule caut cuvant_de_cautat,de exemplu,"A1:H100"
        //ce_caut poate fi: "Adresa",returneaza adresa celula ce contine cuvant_de_cautat
        //"Coloana",returneaza adresa celula ce contine cuvant_de_cautat 
        //"Rand",returneaza adresa celula ce contine cuvant_de_cautat
        //"Valoare",returneaza Value2 al celulei ce contine cuvantul cuvant_de_cautat
        public string Gaseste_adresa_text_in_sheet(string cuvant_de_cautat, string workbook_de_deschis, string sheet_de_deschis,string array_de_cautare,string ce_caut)
        {
            string valoare_gasita = "";
            //string File_name = "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm";
            string File_name = workbook_de_deschis;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;

            try
            {
                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                //daca incerc sa ctivez un sheet activat da eroare

                string nume_sheet_activ = "";
                nume_sheet_activ = oWB.ActiveSheet.Name;
                Debug.WriteLine("oWB.ActiveSheet.Name e:" + nume_sheet_activ);

                if (oWB.ActiveSheet.Name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;
                Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(cuvant_de_cautat, array_de_cautare, oSheet);
                if (oRng != null)
                {
                    //MessageBox.Show("Text found, position is Row:" + oRng.Row + " and column:" + oRng.Column);
                    //MessageBox.Show("Text found, position is:" + oRng.Address);
                    switch (ce_caut)
                    {
                        case "Adresa":
                            valoare_gasita = oRng.Address;
                            break;
                        case "Coloana":
                            valoare_gasita = oRng.Column.ToString();
                            break;
                        case "Rand":
                            valoare_gasita = oRng.Row.ToString();
                            break;
                        case "Valoare":
                        valoare_gasita = oRng.Value2.ToString();
                        break;
                    }
                }
                else
                {
                    MessageBox.Show("Adresa nu e gasita");
                    valoare_gasita = "";
                }

                oWB.Close(false, missing, missing);

                oSheet = null;
                oXL.Quit();

                //MessageBox.Show("Valoarea gasita=" + valoare_gasita);

                return valoare_gasita;
            }
            catch (Exception ex)
            {
                return valoare_gasita;
            }
        }

        //Returneaza ce_caut (randul,adresa,randul,sau coloana) celulei ce contine un cuvant cuvant_de_cautat
        //doar sub celula sub care apare cuvant_sector (daca exista)
        public string Gaseste_informatii_text_in_sector_sheet(string cuvant_de_cautat,string cuvant_sector, string workbook_de_deschis, string sheet_de_deschis, string array_de_cautare, string ce_caut)
        {
            string valoare_gasita = "";
            //string File_name = "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm";
            string File_name = workbook_de_deschis;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            int nr_coloana_cuvant_sector = 0;
            int nr_rand_cuvant_sector=0;
            string coloana_cuvant_sector = "";
            string array_de_cautare_nou = "";

            if (Gaseste_adresa_text_in_sheet(cuvant_sector, workbook_de_deschis, sheet_de_deschis, array_de_cautare, "Rand") != "")
            {
                //"A1:G100"
                //"A19:G100"
                nr_rand_cuvant_sector = Int32.Parse(Gaseste_adresa_text_in_sheet(cuvant_sector, workbook_de_deschis, sheet_de_deschis, array_de_cautare, "Rand"));
                //'A' nu e prea in regula aici dar mno,eventual inlocuit mai tarziu?
                string array_de_eliminat = Operatiuni_String.Returneaza_string_intre_2_charuri(array_de_cautare, 'A', ':');
                int pozitie_inceput_array_de_eliminat = 0;
                array_de_cautare_nou = Operatiuni_String.Elimina_caractere_din_string(array_de_cautare,Operatiuni_String.Returneaza_index_caracter_in_string(array_de_cautare, char.Parse(array_de_eliminat)), array_de_eliminat.Length);
                array_de_cautare_nou = Operatiuni_String.Adauga_un_substring_in_string(array_de_cautare_nou, nr_rand_cuvant_sector.ToString(), 1);

                Debug.WriteLine("Array-ul de cautare nou e:"+ array_de_cautare_nou);

                //nr_coloana_cuvant_sector = Int32.Parse(Gaseste_adresa_text_in_sheet(cuvant_sector, workbook_de_deschis, sheet_de_deschis, array_de_cautare, "Coloana"));
                coloana_cuvant_sector = Returneaza_litera_coloana_dupa_nr(nr_coloana_cuvant_sector);

                try
                {
                    object missing = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing);
                    //daca incerc sa ctivez un sheet activat da eroare

                    string nume_sheet_activ = "";
                    nume_sheet_activ = oWB.ActiveSheet.Name;
                    Debug.WriteLine("oWB.ActiveSheet.Name e:" + nume_sheet_activ);

                    if (oWB.ActiveSheet.Name != sheet_de_deschis)
                    {
                        oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                    }

                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;
                    Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(cuvant_de_cautat, array_de_cautare_nou, oSheet);
                    if (oRng != null)
                    {
                        //MessageBox.Show("Text found, position is Row:" + oRng.Row + " and column:" + oRng.Column);
                        //MessageBox.Show("Text found, position is:" + oRng.Address);
                        switch (ce_caut)
                        {
                            case "Adresa":
                                valoare_gasita = oRng.Address;
                                break;
                            case "Coloana":
                                valoare_gasita = oRng.Column.ToString();
                                break;
                            case "Rand":
                                valoare_gasita = oRng.Row.ToString();
                                break;
                            case "Valoare":
                                valoare_gasita = oRng.Value2.ToString();
                                break;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Adresa nu e gasita");
                        valoare_gasita = "";
                    }

                    oWB.Close(false, missing, missing);

                    oSheet = null;
                    oXL.Quit();

                    //MessageBox.Show("Valoarea gasita=" + valoare_gasita);

                    return valoare_gasita;
                }
                catch (Exception ex)
                {
                    return valoare_gasita;
                }
            }
            return valoare_gasita;
        }

        //nr_rand (mai jos ii 17 da ii numa de proba) Excel.Range oRng = (Excel.Range)oSheet.Range[oSheet.Cells[nr_rand, 2], oSheet.Cells[nr_rand, 19]];
        //nr rand trebuie citit din lista generata de Gaseste_lista_randuri_caractere_inceput astfel incat sa nu plimbam decat pe randurile unde avem grosimea de tabla 8 sau 10 dupa caz
        //.NET transforma datele calendaristice din Excel in string-uri de 5 cifre (nu stiu de ce) -> Toate string-urile ce pot fi convertite in integer de 5 cifre trebuie memorate intr-o lista (eventual si adresele celulelor pt ca o sa ne trebuiasca ca sa gasim pretul cel mai recent mai incolo)
        //Trebuie comparate toate datele din lista si gasita cea mai recenta apoi trebuie sa citim cu o celula mai la stanga ca sa gasim pretul cel mai nou

        //Cauta toate datele in workbook-ul workbook_de_deschis,la sheet-ul sheet_de_deschis,randurile cu nr din Lista_randuri si,returneaza datele din celule gasite si celulele cu o coloana mai inainte sub forma unei liste de obiecte de tip Pret
        //Ordonate dupa data de la cea mai recenta (primul element) la cea mai indepartata
        public List<Pret> Gaseste_lista_preturi_din_lista_randuri(string workbook_de_deschis, string sheet_de_deschis,List <int> Lista_randuri)
        {
            string valoare_gasita = "";
            //string File_name = "S:\\Preturi\\PRETURI MATERIALE\\PRETURI MATERIALE.xlsm";
            string File_name = workbook_de_deschis;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            List<Pret> preturi_gasite = new List<Pret>();

            try
            {
                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;
                //Cells[nr randului,nr coloanei]
                //La nr coloanei: A=1,B=2,C=3,etc

                foreach (int numar_rand in Lista_randuri)
                {
                    oRng = (Excel.Range)oSheet.Range[oSheet.Cells[numar_rand, 2], oSheet.Cells[numar_rand, 19]];
                    //Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(cuvant_de_cautat, array_de_cautare, oSheet);
                    if (oRng != null)
                    {
                        string str = "";
                        string str_offset = "";

                        foreach (Microsoft.Office.Interop.Excel.Range cell in oRng)
                        {
                            //if(cell.v)
                            //{
                            if (!(cell.Value2 == null))
                            {
                                str = cell.Value2.ToString();

                                //Am nevoie de toate astea pt ca,datele de tip Date din excel le transforma in integer de 5 cifre
                                if ((Verificari.Verifica_daca_un_string_e_integer(str) == true) & (str.Length == 5))
                                {
                                    int int_rez = Int32.Parse(str);
                                    Pret pret_gasit = new Pret();
                                    pret_gasit.Data_primire = DateTime.FromOADate(int_rez);
                                    if (!(cell.Offset[0, -1].Value2==null))
                                    {
                                        str_offset = cell.Offset[0, -1].Value2.ToString();
                                        if (Verificari.Verifica_daca_un_string_e_double(str_offset) == true)
                                        {
                                            //cell.Offset[0, -1] 0 offset coloane -1 offset randuri
                                            pret_gasit.Valoare_RON = Double.Parse(str_offset);
                                            //Debug.WriteLine("Valoare rand precedent e:" + str_offset);
                                            //Debug.WriteLine("Valoare rand e:" + DateTime.FromOADate(int_rez));
                                        }
                                        else
                                        {
                                            Debug.WriteLine(str_offset+ " nu e double ");
                                        }
                                    }



                                    preturi_gasite.Add(pret_gasit);
                                    Debug.WriteLine(str + " convertit in data e=" + DateTime.FromOADate(int_rez));
                                    Debug.WriteLine(DateTime.FromOADate(int_rez) + " e data valida=" + Verificari.Verifica_daca_string_e_data(DateTime.FromOADate(int_rez).ToString()));

                                    //Debug.WriteLine("44911 si 44932=" + (DateTime.FromOADate(44911) < DateTime.FromOADate(44932)));
                                }

                                if (Verificari.Verifica_daca_string_e_data2(str) == true)
                                {
                                    var parsedDate = DateTime.Parse(str);
                                    Debug.WriteLine("cell value=" + parsedDate);
                                }
                            }

                            //}
                        }
                    }
                    else
                    {
                        valoare_gasita = "";
                    }
                }

                oWB.Close(false, missing, missing);

                oSheet = null;
                oXL.Quit();

                if (Verificari.Verifica_daca_o_lista_de_preturi_nu_e_goala(preturi_gasite))
                {
                    return preturi_gasite;
                }
                else
                {
                    return null;
                }

                //MessageBox.Show("Valoarea gasita=" + valoare_gasita);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        //Se uita in adresa_workbook,nume_sheet si,pt toate celulele din range-ul range_cautare care incep cu caractere_inceput extrage o lista de preturi (Valoare_RON+data), dupa care le ordoneaza descrescator in functie de data si returneaza pretul cu data cea mai recenta
        public double Returneaza_cel_mai_recent_pret(string adresa_workbook,string nume_sheet,string range_cautare, string caractere_inceput)
        {

            List<Pret> lista_preturi = new List<Pret>();

            if (Verificari.Verifica_daca_o_lista_de_preturi_nu_e_goala(Gaseste_lista_preturi_din_lista_randuri(adresa_workbook, nume_sheet, Gaseste_lista_randuri_caractere_inceput(caractere_inceput, adresa_workbook, nume_sheet, range_cautare))))
            {
                lista_preturi = Gaseste_lista_preturi_din_lista_randuri(adresa_workbook,nume_sheet, Gaseste_lista_randuri_caractere_inceput(caractere_inceput, adresa_workbook, nume_sheet, range_cautare));

                //Pt a sorta descendent
                var lista_ordonata = lista_preturi.OrderByDescending(x => x.Data_primire).ToList();
                //lista_preturi.Sort((a, b) => b.CompareTo(a));

                for (var i = 0; i < lista_ordonata.Count; i++)
                {
                    Debug.WriteLine("Data nr " + i + " " + lista_ordonata[i].Data_primire);
                    Debug.WriteLine("Pretul nr " + i + " [RON] " + lista_ordonata[i].Valoare_RON);
                }

                return lista_ordonata[0].Valoare_RON;
            }
            return 0;
        }

        //Primeste ca string un workbook_de_deschis,sheet-ul in care sa caut,un rand si o coloana si returneaza continutul celulei respective sub forma unui string 
        public string Returneaza_valoarea_de_la_adresa(string workbook_de_deschis, string sheet_de_deschis,int coloana_celula,int rand_celula)
        {
            Microsoft.Office.Interop.Excel.Workbook excel_workbook;

            string rezultat = "";
            //Microsoft.Office.Interop.Excel.Worksheet excel_sheet;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                excel_workbook = excel_app.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                                missing, missing, missing, missing, missing, missing,
                                missing, missing, missing, missing);


                //excel_app.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

                //excel_sheet = (Microsoft.Office.Interop.Excel.Worksheet)excel_workbook.ActiveSheet;

                //Operatiuni_String os = new Operatiuni_String();

                //Ca sa nu dea eroare Math.Round la rotunjire
                //System.Diagnostics.Debug.WriteLine("Pretul gasit " + excel_workbook.Worksheets[sheet_de_deschis].Cells[coloana_celula, rand_celula]);
                if (Operatiuni_String.String_e_Double(excel_workbook.Worksheets[sheet_de_deschis].Cells[rand_celula, coloana_celula].Value2.ToString())==true)
                {
                    rezultat = Math.Round(excel_workbook.Worksheets[sheet_de_deschis].Cells[rand_celula, coloana_celula].Value2, 3).ToString();
                }
                else
                {
                    rezultat = excel_workbook.Worksheets[sheet_de_deschis].Cells[rand_celula, coloana_celula].Value2.ToString();
                }
                
                //False=inchid workbook-ul fara sa salvez
                excel_workbook.Close(false, missing, missing);

                excel_workbook = null;
                excel_app.Quit();

                return rezultat;

            }
            catch
            {
                return "";
            }
        }

        private Microsoft.Office.Interop.Excel.Range GetSpecifiedRange(string matchStr, string range_in_care_caut, Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range currentFind = null;
            Microsoft.Office.Interop.Excel.Range firstFind = null;
            currentFind = objWs.get_Range(range_in_care_caut).Find(matchStr, missing,
                           Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                           Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                           Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            return currentFind;
        }

        //Va crea un excel workbook nou si il va salva cu numele nume_workbook_nou
        public void Creaza_workbook_excel_nou(string nume_workbook_nou)
        {
            //vom adauga nume_workbook_nou+n,pana cand intalnim un nume de fisier care nu exista
            int n = 0;
            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application xl = null;
            _Workbook wb = null;

            // Option 2
            xl = new Microsoft.Office.Interop.Excel.Application();
            xl.SheetsInNewWorkbook = 1;
            xl.Visible = true;
            wb = (_Workbook)(xl.Workbooks.Add(Missing.Value));

            

            wb.SaveAs(Gaseste_path_nefolosit(@"S:\Preturi\PRETURI MATERIALE\Calculator spirala\Calculator_spirala\" + nume_workbook_nou + ".xls"), missing,
    missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
    missing, missing, missing, missing, missing);

        }


        //Metoda ce primeste ca parametru un path(adresa completa a unui fisier),verifica daca exista si,daca da adauga o cifra cand cand gaseste un path cu un fisier care nu exista
        public string Gaseste_path_nefolosit(string path_primit)
        {
            int count = 0;

            string fileNameOnly = Path.GetFileNameWithoutExtension(path_primit);
            string extension = Path.GetExtension(path_primit);
            string path = Path.GetDirectoryName(path_primit);
            string newFullPath = path_primit;

            while (File.Exists(newFullPath))
            {
                string tempFileName = string.Format("{0}({1})", fileNameOnly, count++);
                newFullPath = Path.Combine(path, tempFileName + extension);
            }
            return newFullPath;
        }

        //Incearca sa deschida un obiect de tip Workbook de la adresa filepath_workbook_de_deschis si,daca nu il gaseste,returneaza null
        public Microsoft.Office.Interop.Excel.Workbook Deschide_workbook_dupa_path(string filepath_workbook_de_deschis)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Workbook memo;
            //Microsoft.Office.Interop.Excel.Worksheet oSheet;
            //IsNullOrWhiteSpace=Daca e gol
            //!=Not in C#
            if (!(string.IsNullOrWhiteSpace(filepath_workbook_de_deschis)))
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Open(filepath_workbook_de_deschis, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing);
                    //oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();

                    //Degeaba,tot o sa ramana un Workbook deschis in fundal (il poti vedea in Task Manager)
                    //Nu stiu cum sa rezolv asta
                    memo = oWB;

                    oWB.Close(false, missing, missing);
                    oXL.Quit();

                    return memo;

                }
                catch (Exception ex)
                {
                    return null;
                }
            }

            return null;
        }

        //Insereaza un nr de nr_randuri_de_inserat de randuri incepand de la pozitia nr_rand
        public void Insereaza_randuri_worksheet(string workbook_de_deschis, string sheet_de_deschis, int nr_rand, int nr_randuri_de_inserat)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                // 1. To Delete Entire Row - below rows will shift up
                //Am pus 1,avand in vedere ca se sterge tot randul nu prea conteaza coloana
                Excel.Range cel = oSheet.Cells[nr_rand, 1];



                while (nr_randuri_de_inserat > 0)
                {
                    nr_randuri_de_inserat--;
                    //datele de pe randul nr_rand curent vot fi mutate cu o pozitie mai jos
                    oSheet.Rows[nr_rand].Insert();
                }

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch
            {

            }
        }

        //Schimba Value2 pt celula cu randul nr_rand si coloana nr_col
        public void Schimba_value_celula(string workbook_de_deschis, string sheet_de_deschis, int nr_rand, int nr_col,string valoare_noua)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                oSheet.Cells[nr_rand, nr_col].Value = valoare_noua;

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch
            {

            }
        }

        //Sterge un numar de nr_randuri_de_sters de randuri intr-un fisier excel,incepand de la randul cu nr nr_rand
        public void Sterge_randuri_worksheet(string workbook_de_deschis, string sheet_de_deschis, int nr_rand, int nr_randuri_de_sters)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name!= sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }
                
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                // 1. To Delete Entire Row - below rows will shift up
                //Am pus 1,avand in vedere ca se sterge tot randul nu prea conteaza coloana
                Excel.Range cel = oSheet.Cells[nr_rand, 1];
                while (nr_randuri_de_sters >0)
                {
                    nr_randuri_de_sters--;
                    cel = oSheet.Cells[nr_rand, 1];
                    cel.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                }

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch
            {
                
            }
        }

        //Sterge continutul unor celule aflate intr-un range intre nr_rand1,nr_coloana1 si nr_rand2,nr_coloana2
        public void Sterge_continut_range(string workbook_de_deschis, string sheet_de_deschis, int nr_rand1, int nr_col1, int nr_rand2, int nr_col2)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                // 1. To Delete Entire Row - below rows will shift up
                //Am pus 1,avand in vedere ca se sterge tot randul nu prea conteaza coloana
                Excel.Range cel;

                int rand1;
                int rand2;

                int col1;
                int col2;

                //ne asiguram ca rand1 va fi intodeauna <=rand2
                if (Altele.Max_2_int(nr_rand1, nr_rand2)==nr_rand1)
                {
                    rand1 = nr_rand2;
                    rand2 = nr_rand1;
                }
                else
                {
                    rand1 = nr_rand1;
                    rand2 = nr_rand2;
                }

                if (Altele.Max_2_int(nr_col1, nr_col2) == nr_col1)
                {
                    col1 = nr_col2;
                    col2 = nr_col1;
                }
                else
                {
                    col1 = nr_col1;
                    col2 = nr_col2;
                }

                for (int  i= rand1; i <= rand2; i++)
                {
                    for (int j = col1; j <= col2; j++)
                    {
                        oSheet.Cells[i, j] = "";
                    }
                }

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch
            {

            }
        }

        //Copiaza formatatrea unei celule de la randul si coloana rand_celula_de_copiat si col_celula_de_copiat intr-un range cu unul sau mai multe celule
        public void Copiaza_formatare_celula_in_range(string workbook_de_deschis, string sheet_de_deschis,int rand_celula_de_copiat, int col_celula_de_copiat, int nr_rand1, int nr_col1, int nr_rand2, int nr_col2)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

                //celula cu formatarea care trebuie copiata
                Excel.Range celula_de_copiat;

                int rand1;
                int rand2;

                int col1;
                int col2;

                //ne asiguram ca rand1 va fi intodeauna <=rand2
                if (Altele.Max_2_int(nr_rand1, nr_rand2) == nr_rand1)
                {
                    rand1 = nr_rand2;
                    rand2 = nr_rand1;
                }
                else
                {
                    rand1 = nr_rand1;
                    rand2 = nr_rand2;
                }

                if (Altele.Max_2_int(nr_col1, nr_col2) == nr_col1)
                {
                    col1 = nr_col2;
                    col2 = nr_col1;
                }
                else
                {
                    col1 = nr_col1;
                    col2 = nr_col2;
                }

                celula_de_copiat = oSheet.Cells[rand_celula_de_copiat, col_celula_de_copiat];
                celula_de_copiat.Copy();

                for (int i = rand1; i <= rand2; i++)
                {
                    for (int j = col1; j <= col2; j++)
                    {
                        oSheet.Cells[i, j].PasteSpecial(XlPasteType.xlPasteFormats);
                    }
                }

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch
            {

            }
        }

        //Adauga borders pt toate celulele din workbook_de_deschis,sheet-ul sheet_de_deschis
        //de la celula cu randul si coloana rand_1 si col_1 pana la celula cu rand_2 si col_2
        public void Adauga_chenare_pt_celula_in_range(string workbook_de_deschis, string sheet_de_deschis, int nr_rand_1,int nr_col_1,int nr_rand_2,int nr_col_2)
        {
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();

                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(workbook_de_deschis, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                if (oWB.ActiveSheet.name != sheet_de_deschis)
                {
                    oXL.ActiveWorkbook.Sheets[sheet_de_deschis].Activate();
                }

                int r1 = nr_rand_1;
                int r2 = nr_rand_2;

                //ne asiguram ca rand1 va fi intodeauna <=rand2
                if (Altele.Max_2_int(nr_rand_1, nr_rand_2) == nr_rand_1)
                {
                    r1 = nr_rand_2;
                    r2 = nr_rand_1;
                }
                else
                {
                    r1 = nr_rand_1;
                    r2 = nr_rand_2;
                }

                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;


                Excel.Range beginRange = oSheet.Cells[r1, nr_col_1];
                Excel.Range endRange= oSheet.Cells[r2, nr_col_2];
                oRng = oSheet.Range[beginRange, endRange];

                //celula cu formatarea care trebuie copiata

                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //True=inchid workbook-ul si salvez
                oWB.Close(true, missing, missing);

                oXL = null;
                excel_app.Quit();
            }
            catch (Exception ex)
            {

            }
        }


    }
}
