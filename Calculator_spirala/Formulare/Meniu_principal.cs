using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Calculator_spirala
{

    public partial class Meniu_principal : Form
    {
        public Meniu_principal()
        {
            InitializeComponent();
            System.Diagnostics.Trace.WriteLine("message");
            System.Console.Write("Hello");

        }

        // Ce se intampla cand dau click pe poza cu stalpul central
        private void Poza_stalp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Trace.WriteLine("message");

            Stalp_central stalp_cent = new Stalp_central();
            stalp_cent.ShowDialog(); // Shows Form Stalp central
        }

    }
}
