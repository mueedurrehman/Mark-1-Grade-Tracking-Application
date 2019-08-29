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
    public partial class Welcome3 : MetroFramework.Forms.MetroForm
    {
        public Welcome3()
        {
            InitializeComponent();
        }

        private void Welcome3_Load(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                metroTile1.Enabled = true;
            }
            else
            {
                metroTile1.Enabled = false;
            }
        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {

        }
    }
}
