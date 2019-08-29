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


    public partial class Mainform1 : MetroFramework.Forms.MetroForm
    {
        public Mainform1()
        {
            InitializeComponent();
        }
        public static List<List<string>> semesterlist = new List<List<string>>();
        public static List<List<string>> CLOs = new List<List<string>>();
        public static List<double> Studentids = new List<double>();
        public static List<string> studentnames = new List<string>();
        public static List<List<string>> semesterlist1 = new List<List<string>>();
        public static List<List<List<List<List<double>>>>> studentdata = new List<List<List<List<List<double>>>>>();
        public static string pathString2;
        public static int NumberOfSemesters = 8;
        public static int[,] PLOWeightageCLO;
        public static int[,] PLOWeightageSem;
        public static int combined = 1;
        public static int outputcontrol;
        public static string input4;
        public static string pathString;
        public static int promptindex;
        private void Mainform1_Load(object sender, EventArgs e)
        {
            
            
            outputcontrol = 0;
            Welcome2 wel = new Welcome2();
            wel.ShowDialog();

            Welcome3 wel1 = new Welcome3();
            wel1.ShowDialog();

            //MessageBox.Show("The user is required to follow these instructions:\n\n    1.   The setup file must be created if it hasnt been created already.\n    2.   Required information including Student Names/ID and Number \n          of CLOs for each and every course must be entered.\n    3.   PLO associated with each and every CLO must be initialised.\n    4.   The marks entered must be out of 100.", "Instructions");
            Filepath filepathandsetupyesno = new Filepath();
            filepathandsetupyesno.ShowDialog();
            input4 = Filepath.path;


            //DialogResult dialogResult1 = MessageBox.Show("Would you like the program to initiate setup procedure?", "Marksheet Handler", MessageBoxButtons.YesNo);
            if (Filepath.metroRB1 == 1)
            {
                Setup set = new Setup();
                set.ShowDialog();
                string input3 = input4 + "\\Setup file.xls";
               
                string folderName = input4;
                pathString = System.IO.Path.Combine(folderName, "Marksheet Handling Program");

                System.IO.Directory.CreateDirectory(pathString);

               
                outputcontrol = 1;
                ProgressBar progress1 = new ProgressBar();
                progress1.ShowDialog();


                promptindex = 0;
                DialogResult Programconfirm;
                System.Diagnostics.Process.Start("explorer.exe", pathString);
                do
                {
                    
                    prompts prompt1 = new prompts();
                    prompt1.ShowDialog();
                    Programconfirm = MetroFramework.MetroMessageBox.Show(this,"Are you sure you have entered the required description and want to proceed?", "Confirmation", MessageBoxButtons.YesNo,MessageBoxIcon.Information,150);
                } while (Programconfirm == DialogResult.No);

                semesterlist1.Clear();

            }

            outputcontrol = 2;
            ProgressBar progress2 = new ProgressBar();
            progress2.ShowDialog();
       
            string pathstring = input4 + "\\Marksheet Handling Program";
            pathString2 = System.IO.Path.Combine(pathstring, "Marksheets");
            System.IO.Directory.CreateDirectory(pathString2);
            if (Filepath.metroCB1 == 0)
            {
                outputcontrol = 3;
                ProgressBar progress3 = new ProgressBar();
                progress3.ShowDialog();

                promptindex = 1;
                DialogResult semesterconfirm;
                System.Diagnostics.Process.Start("explorer.exe", pathString2);
                do
                {

                    
                    prompts prompt2 = new prompts();
                    prompt2.ShowDialog();
                    semesterconfirm = MetroFramework.MetroMessageBox.Show(this,"Are you sure you have entered the marks and want to proceed?", "Confirmation", MessageBoxButtons.YesNo,MessageBoxIcon.Information,150);
                } while (semesterconfirm == DialogResult.No);
            }

            outputcontrol = 4;
            ProgressBar progress4 = new ProgressBar();
            progress4.ShowDialog();
       
            StudentNamesCount = studentnames.Count;
        }


        public static int StudentNamesCount;

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
           
                OutputExcel.outputagain = 8;
                metroTile1.Enabled = true;
            
        }

        private void metroRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
           
                OutputExcel.outputagain = 3;
                metroTile1.Enabled = true;
            
        }

        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            
                OutputExcel.outputagain = 1;
                metroTile1.Enabled = true;
           
        }

        private void metroRadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            
                OutputExcel.outputagain = -1;
                OutputExcel.outputagain2 = 2;
                metroTile1.Enabled = true;
          
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {

            OutputExcel.outputagain = -1;
            OutputExcel.outputagain2 = -1;
            Close();
        }

    }

}
