using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EC_102_Project
{
    public partial class CourseAdd : MetroFramework.Forms.MetroForm
    {
        public CourseAdd()
        {
            InitializeComponent();
        }

        private void CourseAdd_Load(object sender, EventArgs e)
        {
            for (int j = 0; j < Mainform1.semesterlist1.Count; j++)
            {
                metroComboBox1.Items.Add(j + 1);
            }
        }

    
        private void metroTile1_Click(object sender, EventArgs e)
        {

            int input8 =metroComboBox1.SelectedIndex;
            string input7 = metroTextBox2.Text;
            Mainform1.semesterlist1[input8].Add(input7);
           
            Close();
        }
    }
}
