using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using System.Threading;
using System.Linq;
using System.Data;

namespace EC_102_Project
{
    public partial class Setup :MetroFramework.Forms.MetroForm
    {
        public Setup()
        {
            InitializeComponent();
        }
        public static int coursecheck = 0;
        private void Setup_Load(object sender, EventArgs e)
        {
         
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (metroRadioButton1.Checked == true)
            { 
                coursecheck = 0;
                metroTextBox1.Enabled = false;
                metroButton2.Enabled = false;
            
                timer1.Start();
                timer2.Start();
            }
        }

        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (metroRadioButton2.Checked == true)
            {
                coursecheck = 1;
                metroTextBox1.Enabled = true;
                metroButton1.Enabled = false;
                metroButton2.Enabled = true;
                metroButton3.Enabled = false;
                metroGrid1.Columns.Clear();
                metroGrid1.Rows.Clear();
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            CourseAdd course = new CourseAdd();
            course.ShowDialog();
            timer2.Start();

        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            CourseRemove course1 = new CourseRemove();
            course1.ShowDialog();
            timer2.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            if (coursecheck == 0)
            {
                Mainform1.semesterlist1.Clear();
                Mainform1.NumberOfSemesters = 8;
                #region Addition/Removal in Default Courses
                Mainform1.semesterlist1.Add(new List<string>(0));
                Mainform1.semesterlist1.Add(new List<string>(1));
                Mainform1.semesterlist1.Add(new List<string>(2));
                Mainform1.semesterlist1.Add(new List<string>(3));
                Mainform1.semesterlist1.Add(new List<string>(4));
                Mainform1.semesterlist1.Add(new List<string>(5));
                Mainform1.semesterlist1.Add(new List<string>(6));
                Mainform1.semesterlist1.Add(new List<string>(7));

                Mainform1.semesterlist1[0].Add("Semester 1");
                Mainform1.semesterlist1[1].Add("Semester 2");
                Mainform1.semesterlist1[2].Add("Semester 3");
                Mainform1.semesterlist1[3].Add("Semester 4");
                Mainform1.semesterlist1[4].Add("Semester 5");
                Mainform1.semesterlist1[5].Add("Semester 6");
                Mainform1.semesterlist1[6].Add("Semester 7");
                Mainform1.semesterlist1[7].Add("Semester 8");


                Mainform1.semesterlist1[0].Add("PHY - 102");
                Mainform1.semesterlist1[0].Add("MATH - 105");
                Mainform1.semesterlist1[0].Add("EC - 102");
                Mainform1.semesterlist1[0].Add("HU - 100");
                Mainform1.semesterlist1[0].Add("HU - 101");
                Mainform1.semesterlist1[0].Add("ME - 110");
                Mainform1.semesterlist1[0].Add("ME - 121");

                Mainform1.semesterlist1[1].Add("ME - 111");
                Mainform1.semesterlist1[1].Add("ME - 112");
                Mainform1.semesterlist1[1].Add("ME - 130");
                Mainform1.semesterlist1[1].Add("MATH - 121");
                Mainform1.semesterlist1[1].Add("CH - 101");
                Mainform1.semesterlist1[1].Add("HU - 107");
                Mainform1.semesterlist1[1].Add("HU - 109");

                Mainform1.semesterlist1[2].Add("MATH - 241");
                Mainform1.semesterlist1[2].Add("ME - 213");
                Mainform1.semesterlist1[2].Add("ME - 220");
                Mainform1.semesterlist1[2].Add("ME - 230");
                Mainform1.semesterlist1[2].Add("ME - 236");

                Mainform1.semesterlist1[3].Add("MATH - 231");
                Mainform1.semesterlist1[3].Add("EE - 103");
                Mainform1.semesterlist1[3].Add("HU - 222");
                Mainform1.semesterlist1[3].Add("HU - 212");
                Mainform1.semesterlist1[3].Add("ME - 211");
                Mainform1.semesterlist1[3].Add("ME - 235");

                Mainform1.semesterlist1[4].Add("MATH - 361");
                Mainform1.semesterlist1[4].Add("ME - 310");
                Mainform1.semesterlist1[4].Add("ME - 221");
                Mainform1.semesterlist1[4].Add("ME - 312");
                Mainform1.semesterlist1[4].Add("ME - 323");
                Mainform1.semesterlist1[4].Add("ME - 325");
                Mainform1.semesterlist1[4].Add("EE - 212");

                Mainform1.semesterlist1[5].Add("MATH - 351");
                Mainform1.semesterlist1[5].Add("ME - 420");
                Mainform1.semesterlist1[5].Add("ME - 311");
                Mainform1.semesterlist1[5].Add("ME - 315");
                Mainform1.semesterlist1[5].Add("ME - 330");
                Mainform1.semesterlist1[5].Add("ME - 331");
                Mainform1.semesterlist1[5].Add("ME - 332");

                Mainform1.semesterlist1[6].Add("ME - 314");
                Mainform1.semesterlist1[6].Add("ME - 421");
                Mainform1.semesterlist1[6].Add("ME - 410");
                Mainform1.semesterlist1[6].Add("ME - 448");
                Mainform1.semesterlist1[6].Add("XX - 4XX");
                Mainform1.semesterlist1[6].Add("ME - 499");
                Mainform1.semesterlist1[6].Add("CSL - 401");

                Mainform1.semesterlist1[7].Add("MGT - 271");
                Mainform1.semesterlist1[7].Add("ME - 498");
                Mainform1.semesterlist1[7].Add("XX - 4XX2");
                Mainform1.semesterlist1[7].Add("XX - 4XX3");
                Mainform1.semesterlist1[7].Add("ME - 499");




                #endregion
            }
            else
            {
                Mainform1.semesterlist1.Clear();
                int numsem;
                if (string.IsNullOrWhiteSpace(metroTextBox1.Text))
                { 
                    numsem = 0;
                }
                else
                {
                    numsem = Convert.ToInt32(metroTextBox1.Text);
                   
                }
                    Mainform1.NumberOfSemesters = numsem;
                
                for (int j=0;j<numsem;j++)
                {
                    Mainform1.semesterlist1.Add(new List<string>(j));
                    Mainform1.semesterlist1[j].Add("Semester "+Convert.ToString(j+1));
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Stop();

            metroGrid1.Columns.Clear();
            metroGrid1.Rows.Clear();
            
            DataGridViewRow row;
            for (int j = 0; j < Mainform1.semesterlist1.Count; j++)
            {
                DataGridViewColumn col = new DataGridViewColumn();
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                col.CellTemplate = cell;
                col.HeaderText = Mainform1.semesterlist1[j][0];
                metroGrid1.Columns.Add(col);

            }
            int maxcount = 0;
            for (int z = 0; z < Mainform1.semesterlist1.Count; z++)
            {

                if (Mainform1.semesterlist1[z].Count > maxcount)
                {
                    maxcount = Mainform1.semesterlist1[z].Count;
                   
                }

            }
            if (maxcount<2)
            {
                metroButton3.Visible = false;
                metroButton3.Enabled = false;
            }
            else
            {
                metroButton3.Visible = true;
                metroButton3.Enabled = true;
            }
            if (Mainform1.NumberOfSemesters==0)
            {
                metroButton1.Visible = false;
                metroButton1.Enabled = false;
            }
            else
            {
                metroButton1.Visible = true;
                metroButton1.Enabled = true;
            }

            for (int i = 1; i < maxcount; i++)
            {
                row = new DataGridViewRow();
                row.CreateCells(metroGrid1);
                for (int j = 0; j < Mainform1.semesterlist1.Count; j++)
                {
                    if (i < Mainform1.semesterlist1[j].Count)
                    {
                        row.Cells[j].Value = Mainform1.semesterlist1[j][i];
                    }
                    else
                    {
                        row.Cells[j].Value = "";
                    }
                }
                metroGrid1.Rows.Add(row);
                metroGrid1.Update();

            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            timer1.Start();
            Thread.Sleep(200);
            timer2.Start();
        }
    }
}
