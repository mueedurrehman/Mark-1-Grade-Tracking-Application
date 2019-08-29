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
    public partial class Filepath : MetroFramework.Forms.MetroForm
    {
        public Filepath()
        {
            InitializeComponent();
        }

        public static int metroRB1=1, metroRB2=0, metroCB1=0;
        public static string path = "";
       

        private void Filepath_Load(object sender, EventArgs e)
        {
          

        }

       
        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            metroRB1 = 1;
            metroRB2 = 0;
            metroCheckBox1.Enabled = false;
            metroCB1 = 0;
            metroCheckBox1.Checked = false;
        }

        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            metroRB2 = 1;
            metroRB1 = 0;
            metroCheckBox1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (metroRB1 == 1 && metroRB2 == 0)
            {
                if (System.IO.Directory.Exists(metroTextBox1.Text))
                {
                    pictureBox1.Visible = false;
                    metroLabel3.Visible = false;
                    metroTile1.Enabled = true;
                }
                else
                {
                    pictureBox1.Visible = true;
                    metroLabel3.Visible = true;
                    metroTile1.Enabled = false;
                }
            }
            else if (metroRB2 == 1 && metroRB1 == 0)
            {
                string path1 = metroTextBox1.Text + "\\Marksheet Handling Program\\Setup File.xls";
                if (System.IO.File.Exists(path1))
                {
                    pictureBox1.Visible = false;
                    metroLabel3.Visible = false;
                    metroTile1.Enabled = true;
                 
                }
                else
                {
                    pictureBox1.Visible = true;
                    metroLabel3.Visible = true;
                    metroTile1.Enabled = false;
                }
            }
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {
            timer1.Start();
        }
        private void metroCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            metroCB1 = 1;
        }
        private void metroTile1_Click(object sender, EventArgs e)
        {
            path = metroTextBox1.Text;
            Close();
        }
    }
}
