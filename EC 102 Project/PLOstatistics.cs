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
    public partial class PLOstatistics : MetroFramework.Forms.MetroForm
    {
        public PLOstatistics()
        {
            InitializeComponent();
        }

        public static List<int> semesnumform= new List<int>();
        public static List<int> PLOnumform= new List<int>();

        private void PLOstatistics_Load(object sender, EventArgs e)
        {
            DataTable dtEmp = new DataTable();
            // add column to datatable  
            dtEmp.Columns.Add(" ", typeof(bool));
            dtEmp.Columns.Add("PLOs", typeof(string));

          
            for (int i = 0; i < 12; i++)
            {
                dtEmp.Rows.Add(false, Convert.ToString(i+1));
               
            }
            metroGrid1.DataSource = dtEmp;
            metroGrid1.Columns[0].Width = 30;
            metroGrid1.Columns[1].Width = 50;

            DataTable dtEmp1 = new DataTable();
            // add column to datatable  
            dtEmp1.Columns.Add(" ", typeof(bool));
            dtEmp1.Columns.Add("Semesters", typeof(string));
            

          
            for (int i = 0; i < Mainform1.NumberOfSemesters; i++)
            {
                dtEmp1.Rows.Add(false, Convert.ToString(i + 1));

            }
            metroGrid2.DataSource = dtEmp1;
            metroGrid2.Columns[0].Width = 30;
            metroGrid2.Columns[1].Width = 100;






        }

        private void metroTile1_Click(object sender, EventArgs e)
        {


            Close();
        }
        
        private void metroGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in metroGrid1.Rows)
            {
                if ((bool)row.Cells[0].Value == true)
                {
                    
                    PLOnumform.Add(row.Index + 1);
                }

            }
            foreach (DataGridViewRow row in metroGrid2.Rows)
            {
                if ((bool)row.Cells[0].Value == true)
                {
                    
                    semesnumform.Add(row.Index + 1);
                }

            }
            string text1="";
            for (int PLOIndex = 0; PLOIndex < PLOnumform.Count; PLOIndex++)
            {
                for (int semIndex = 0; semIndex < semesnumform.Count; semIndex++)
                {
                    int actualSem = semesnumform[semIndex] - 1;
                    int actualPLO = PLOnumform[PLOIndex] - 1;
                    var statistic = StudentManager.PLOStatistics[actualPLO].PLOSemesterStatistics[actualSem];

                    text1 = text1 + "PLO " + Convert.ToString(actualPLO +1) + " Statistics for Semester " + Convert.ToString(semesnumform[semIndex]) + ":\nThe Average CLO Marks are: " + Convert.ToString(Math.Round(statistic.AverageCLOMarks,3)) + "\nVariance in Average CLO Marks is: " + Convert.ToString(Math.Round(statistic.Variance,3)) + "\nThe Standard Deviation in Average CLO Marks is: " + Convert.ToString(Math.Round(statistic.StandardDeviation,3)) + "\n\n";
                      
                }
            }
            metroLabel2.Visible = true;
            metroPanel1.Visible = true;
            metroLabel3.Text = text1;
            


        }
    }
}
