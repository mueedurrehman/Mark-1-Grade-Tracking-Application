using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using System.Threading;

namespace EC_102_Project
{






    public partial class OutputExcel : MetroFramework.Forms.MetroForm
    {
        public OutputExcel()
        {
            InitializeComponent();
        }
        public int progresspercentageoutput { get; set; }
        public static int outputagain = 0, outputagain2 = 0;
        public static int progress=1;

        private void OutputExcel_Load(object sender1, EventArgs e1)
        {
            if (Mainform1.outputcontrol == 5)
            {

                ProgressBar progress = new ProgressBar();
                progress.ShowDialog();

            }
            else if (Mainform1.outputcontrol == 10)
            {
                PLOstatistics plo1 = new PLOstatistics();
                plo1.ShowDialog();
                metroLabel1.Text = "Please choose what you want to do next.";
            }
            else if (Mainform1.outputcontrol == 11)
            {
                indiStudent stu = new indiStudent();
                stu.ShowDialog();
                metroLabel1.Text = "Please choose what you want to do next.";
            }

            else { }

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }
        

        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
           
            outputagain = 8;
            metroTile1.Enabled = true;
        }

        private void metroRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            outputagain = 3;
            metroTile1.Enabled = true;
        }

        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            outputagain = 1;
            metroTile1.Enabled = true;

        }
        
        

        private void metroButton1_Click(object sender, EventArgs e)
        {
            outputagain = -1;
            outputagain2 = -1;
            Close();
        }


        private void metroRadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            outputagain = -1;
            outputagain2 = 2;
            metroTile1.Enabled = true;
        }
    }
}
