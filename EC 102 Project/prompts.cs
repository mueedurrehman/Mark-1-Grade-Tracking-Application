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
    public partial class prompts : MetroFramework.Forms.MetroForm
    {
        public prompts()
        {
            InitializeComponent();
        }

        private void prompts_Load(object sender, EventArgs e)
        {
            if (Mainform1.promptindex == 0)
            {
                metroLabel1.Text = "Setup File has been generated. Please enter the required information\n and then check the checkboxes before you proceed.";
                metroCheckBox1.Text = "I have entered the Student Names/IDs.";
                metroCheckBox2.Text = "I have entered the No. of CLOs for each course.";


            }
            else
            {
                // Setup File has been generated. Please enter the following information 
                //before you proceed.Make selections to confirm and then click Proceed.
                metroLabel1.Text = "Seperate marksheets for every semester have been generated. Please\n enter the following information, select to confirm and then click Proceed.";
                metroCheckBox1.Text = "I have entered the marks.";
                metroCheckBox2.Visible = false;
                metroCheckBox2.Enabled = false;
            }
       }

        private void metroCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (Mainform1.promptindex == 0)
            {
               if(metroCheckBox2.Checked==true&& metroCheckBox1.Checked==true)
                {
                    metroTile1.Enabled = true;
                }
                else
                {
                    metroTile1.Enabled = false;
                }

            }
            else
            {
                if ( metroCheckBox1.Checked == true)
                {
                    metroTile1.Enabled = true;
                }
                else
                {
                    metroTile1.Enabled = false;
                }
            }
        }

        private void metroCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (Mainform1.promptindex == 0)
            {
                if (metroCheckBox2.Checked == true && metroCheckBox1.Checked == true)
                {
                    metroTile1.Enabled = true;
                }
                else
                {
                    metroTile1.Enabled = false;
                }

            }
            else
            {
                if (metroCheckBox1.Checked == true)
                {
                    metroTile1.Enabled = true;
                }
                else
                {
                    metroTile1.Enabled = false;
                }
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
