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
    public partial class CourseRemove : MetroFramework.Forms.MetroForm
    {
        public CourseRemove()
        {
            InitializeComponent();
        }

        private void CourseRemove_Load(object sender, EventArgs e)
        {
            for (int j = 0; j < Mainform1.semesterlist1.Count; j++)
            {
                if (Mainform1.semesterlist1[j].Count > 1)
                { metroComboBox1.Items.Add(Convert.ToString(j + 1)); }
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            string input9 = metroTextBox2.Text;;
            int input10 = Convert.ToInt32(metroComboBox1.Text)-1;
            
            
            for (int j=1;j<Mainform1.semesterlist1[input10].Count;j++)
            { 
                        if (input9 == Mainform1.semesterlist1[input10][j] || Regex.Match(input9, @"\d+") == Regex.Match(Mainform1.semesterlist1[input10][j], @"\d+"))
                        {
                            Mainform1.semesterlist1[input10].RemoveAt(j);
                        }
            }

            Close();
                
            
        }

    }
}
