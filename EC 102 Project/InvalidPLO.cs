using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EC_102_Project
{
    public partial class InvalidPLO : MetroFramework.Forms.MetroForm
    {
        public InvalidPLO()
        {
            InitializeComponent();
        }

        public static string InvalidPLOcorrect;

        private void Form1_Load(object sender, EventArgs e)
        {
            if (ProgressBar.InvalidPLOindex == 1)
            {
                metroLabel1.Text = "Oops! The PLO associated with CLO " + ProgressBar.InvalidPLOCLO + " of Course " + ProgressBar.InvalidPLOcourse + " in " + ProgressBar.InvalidPLOsem + "\nwas found to be out of the predefined range of PLOs.";
            }
            else
            {
                metroLabel1.Text = "Oops! The PLO associated with CLO " + ProgressBar.InvalidPLOCLO + " of Course " + ProgressBar.InvalidPLOcourse + " in " + ProgressBar.InvalidPLOsem + "\nhas not been specified.";
            }

           
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            InvalidPLOcorrect = comboBox1.Text;
            Close();
        }

    }
}
