using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
using System.Threading;
using System.ComponentModel;

namespace EC_102_Project
{
    public partial class ProgressBar : MetroFramework.Forms.MetroForm
    {
        public ProgressBar()
        {
            InitializeComponent();
        }
       
        public static int outputagain = 0, outputagain2 = 0, combobox = 0;
        public static int progress = 1;
        public static int InvalidPLOindex,InvalidPLOCLO;
        public static string InvalidPLOcourse,InvalidPLOsem ;
                                        


        private void timer1_Tick_1(object sender, EventArgs e)
        {
            timer1.Stop();
            if(Mainform1.outputcontrol==1)
            {
                #region Setup File creation



                Excel.Application Startup = new Excel.Application();

                if (Startup == null)
                {
                    MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return;
                }

                Startup.Visible = false; // To avoid error hresult 0x800ac472
                Excel.Workbook startup = Startup.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Startup.DisplayAlerts = false;
                Excel.Worksheet Prodes = (Excel.Worksheet)startup.Worksheets[1];

                Prodes.Name = "Program Description";
                Startup.Sheets[1].Range[Startup.Sheets[1].Cells[1, 1], Startup.Sheets[1].Cells[90, 14]].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Startup.Sheets[1].Range[Startup.Sheets[1].Cells[1, 1], Startup.Sheets[1].Cells[90, 14]].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                Prodes.Cells[1, 1].EntireColumn.columnwidth = 15;
                Prodes.Cells[1, 2].EntireColumn.columnwidth = 15;
                Prodes.Cells[1, 3].EntireColumn.columnwidth = 15;
                Prodes.Cells[1, 1].EntireRow.Font.Bold = true;
                Prodes.Cells[1, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                Prodes.Cells[1, 2].Interior.Color = System.Drawing.Color.NavajoWhite;
                Prodes.Cells[1, 3].Interior.Color = System.Drawing.Color.NavajoWhite;
                Prodes.Cells[1, 6].Interior.Color = System.Drawing.Color.NavajoWhite;
                Prodes.Cells[1, 8].Interior.Color = System.Drawing.Color.NavajoWhite;

                // ROW 1
                Prodes.Cells[1, 1].Value2 = "Semester";
                Prodes.Cells[1, 2].Value2 = "Course Code";
                Prodes.Cells[1, 3].Value2 = "No. of CLOs";
                Prodes.Range[Prodes.Cells[1, 6], Prodes.Cells[1, 7]].Merge(); Prodes.Cells[1, 6].Value2 = "PLO";
                Prodes.Range[Prodes.Cells[1, 8], Prodes.Cells[1, 14]].Merge(); Prodes.Cells[1, 8].Value2 = "Description";

                // Semester numbers
                int count = 2;
                for (int i = 0; i < Mainform1.NumberOfSemesters; i++)
                {
                    int count1 = count + Mainform1.semesterlist1[i].Count - 2;
                    Prodes.Range[Prodes.Cells[count, 1], Prodes.Cells[count1, 1]].Merge();

                    Prodes.Cells[count, 1].Value2 = Mainform1.semesterlist1[i][0];
                    count = count1 + 1;

                }
                Random random = new Random();
                int randomNumber = random.Next(24, 44);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Creating File... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();

                // Course Codes
                count = 2;
                for (int i = 0; i < Mainform1.NumberOfSemesters; i++)
                {
                    for (int j = 1; j <= Mainform1.semesterlist1[i].Count - 1; j++)
                    {
                        if (j < Mainform1.semesterlist1[i].Count)
                        {
                            Prodes.Cells[count, 2].Value2 = Mainform1.semesterlist1[i][j];
                            count++;
                        }
                        else { j = Mainform1.semesterlist1[i].Count; }
                    }
                }
                random = new Random();
                randomNumber = random.Next(64, 84);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Writing Data... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();

                //CLO column
                for (int n = 2; n <= Mainform1.semesterlist1.Count - 8; n++)
                {
                    Prodes.Cells[n, 3].Value2 = 0;
                }
                random = new Random();
                randomNumber = random.Next(84, 94);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Still working... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();


                // FOR PLOs
                Prodes.Range[Prodes.Cells[2, 6], Prodes.Cells[8, 7]].Merge(); Prodes.Cells[2, 6].Value2 = "PLO 1";
                Prodes.Range[Prodes.Cells[2, 8], Prodes.Cells[8, 14]].Merge(); Prodes.Cells[2, 8].Value2 = "Engineering Knowledge. \nAn ability to apply knowledge of mathematics, science, engineering fundamentals and an engineering specialization to the solution of complex engineering problems.";
                Prodes.Range[Prodes.Cells[9, 6], Prodes.Cells[15, 7]].Merge(); Prodes.Cells[9, 6].Value2 = "PLO 2";
                Prodes.Range[Prodes.Cells[9, 8], Prodes.Cells[15, 14]].Merge(); Prodes.Cells[9, 8].Value2 = "Problem Analysis. \nAn ability to identify, formulate, research literature, and analyse complex engineering problems reaching substantiated conclusions using first principles of mathematics, natural sciences and engineering sciences.";
                Prodes.Range[Prodes.Cells[16, 6], Prodes.Cells[22, 7]].Merge(); Prodes.Cells[16, 6].Value2 = "PLO 3";
                Prodes.Range[Prodes.Cells[16, 8], Prodes.Cells[22, 14]].Merge(); Prodes.Cells[16, 8].Value2 = "Design / Development of Solutions. \nAn ability to design solutions for complex engineering problems and design systems, components or processes that meet specified needs with appropriate consideration for public health and safety, cultural, societal, and environmental considerations.";
                Prodes.Range[Prodes.Cells[23, 6], Prodes.Cells[29, 7]].Merge(); Prodes.Cells[23, 6].Value2 = "PLO 4";
                Prodes.Range[Prodes.Cells[23, 8], Prodes.Cells[29, 14]].Merge(); Prodes.Cells[23, 8].Value2 = "Investigation. \nAn ability to investigate complex engineering problems in a methodical way including literature survey, design and conduct of experiments, analysis and interpretation of experimental data, and synthesis of information to derive valid conclusions.";
                Prodes.Range[Prodes.Cells[30, 6], Prodes.Cells[36, 7]].Merge(); Prodes.Cells[30, 6].Value2 = "PLO 5";
                Prodes.Range[Prodes.Cells[30, 8], Prodes.Cells[36, 14]].Merge(); Prodes.Cells[30, 8].Value2 = "Modern Tool Usage. \nAn ability to create, select and apply appropriate techniques, resources, and modern engineering and IT tools, including prediction and modelling, to complex engineering activities, with an understanding of the limitations.";
                Prodes.Range[Prodes.Cells[37, 6], Prodes.Cells[43, 7]].Merge(); Prodes.Cells[37, 6].Value2 = "PLO 6";
                Prodes.Range[Prodes.Cells[37, 8], Prodes.Cells[43, 14]].Merge(); Prodes.Cells[37, 8].Value2 = "The Engineer and Society. \nAn ability to apply reasoning informed by contextual knowledge to assess societal, health, safety, legal and cultural issues and the consequent responsibilities relevant to professional engineering practice and solution to complex engineering problems.";
                Prodes.Range[Prodes.Cells[44, 6], Prodes.Cells[50, 7]].Merge(); Prodes.Cells[44, 6].Value2 = "PLO 7";
                Prodes.Range[Prodes.Cells[44, 8], Prodes.Cells[50, 14]].Merge(); Prodes.Cells[44, 8].Value2 = "Environment and Sustainability. \nAn ability to understand the impact of professional engineering solutions in societal and environmental contexts and demonstrate knowledge of and need for sustainable development.";
                Prodes.Range[Prodes.Cells[51, 6], Prodes.Cells[57, 7]].Merge(); Prodes.Cells[51, 6].Value2 = "PLO 8";
                Prodes.Range[Prodes.Cells[51, 8], Prodes.Cells[57, 14]].Merge(); Prodes.Cells[51, 8].Value2 = "Ethics. \nApply ethical principles and commit to professional ethics and responsibilities and norms of engineering practice.";
                Prodes.Range[Prodes.Cells[58, 6], Prodes.Cells[64, 7]].Merge(); Prodes.Cells[58, 6].Value2 = "PLO 9";
                Prodes.Range[Prodes.Cells[58, 8], Prodes.Cells[64, 14]].Merge(); Prodes.Cells[58, 8].Value2 = "Individual and Teamwork. \nAn ability to work effectively, as an individual or in a team, on multifaceted and / or multidisciplinary settings.";
                Prodes.Range[Prodes.Cells[65, 6], Prodes.Cells[71, 7]].Merge(); Prodes.Cells[65, 6].Value2 = "PLO 10";
                Prodes.Range[Prodes.Cells[65, 8], Prodes.Cells[71, 14]].Merge(); Prodes.Cells[65, 8].Value2 = "Communication. \nAn ability to communicate effectively, orally as well as in writing, on complex engineering activities with the engineering community and with society at large, such as being able to comprehend and write effective reports and design documentation, make effective presentations, and give and receive clear instructions.";
                Prodes.Range[Prodes.Cells[72, 6], Prodes.Cells[78, 7]].Merge(); Prodes.Cells[72, 6].Value2 = "PLO 11";
                Prodes.Range[Prodes.Cells[72, 8], Prodes.Cells[78, 14]].Merge(); Prodes.Cells[72, 8].Value2 = "Project Management. \nAn ability to demonstrate management skills and apply engineering principles to one’s own work, as a member and/ or leader in a team, to manage projects in a multidisciplinary environment.";
                Prodes.Range[Prodes.Cells[79, 6], Prodes.Cells[85, 7]].Merge(); Prodes.Cells[79, 6].Value2 = "PLO 12";
                Prodes.Range[Prodes.Cells[79, 8], Prodes.Cells[85, 14]].Merge(); Prodes.Cells[79, 8].Value2 = " Lifelong Learning. \nAn ability to recognize importance of, and pursue lifelong learning in the broader context of innovation and technological developments.";
                random = new Random();
                randomNumber = random.Next(94, 97);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Almost Done... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();



                //NEXT SHEET *******************
                Excel.Worksheet Studentid = (Excel.Worksheet)startup.Sheets.Add();
                Studentid.Cells[1, 2].EntireColumn.columnwidth = 15;
                Studentid.Cells[1, 1].EntireColumn.columnwidth = 15;
                Studentid.Name = "Student Names";
                for (int j = 1; j <= 250; j++)
                {
                    Studentid.Cells[j, 1].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Studentid.Cells[j, 2].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                Studentid.Cells[1, 1].EntireRow.Font.Bold = true;
                Studentid.Cells[1, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                Studentid.Cells[1, 2].Interior.Color = System.Drawing.Color.NavajoWhite;



                if (Studentid == null)
                {
                    MessageBox.Show("Worksheet could not be created. Check that your office installation and project references are correct.");
                }
                //Student Names table.
                Studentid.Cells[1, 1].Value2 = "Student ID";
                Studentid.Cells[1, 2].Value2 = "Student Name";
                
                metroProgressBar1.Value = 100;
                lblProcess.Text = "Done! " + 100 + " %";
                metroProgressBar1.Update();
                lblProcess.Update();

                string startupfilesave = Mainform1.pathString + "\\Setup File.xls";
                startup.SaveAs(@startupfilesave, Excel.XlFileFormat.xlWorkbookNormal, null, null, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);

                startup.Close(true, null, null);
                Startup.Quit();
                Marshal.ReleaseComObject(Prodes);
                Marshal.ReleaseComObject(Studentid);
                Marshal.ReleaseComObject(startup);
                Marshal.ReleaseComObject(Startup);
                #endregion
            }
            else if (Mainform1.outputcontrol==2)
            {
                #region Read from Course description
                Excel.Application Startup1;
                Excel.Workbook startup1;
                Excel.Worksheet Studentid1;
                Excel.Range testrange;


                int rCnt;
                int cCnt;
                int rw = 0;
                string s, x;
                double u;
                string input11 = Mainform1.input4 + "\\Marksheet Handling Program\\Setup File.xls";
                Startup1 = new Excel.Application();
                startup1 = Startup1.Workbooks.Open(@input11, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Studentid1 = (Excel.Worksheet)startup1.Worksheets.get_Item("Student Names");
                Startup1.Visible = false;
                Startup1.DisplayAlerts = false;
                testrange = Studentid1.UsedRange;
                rw = testrange.Rows.Count;



                for (rCnt = 1; rCnt < rw; rCnt++)
                {

                    for (cCnt = 1; cCnt <= 2; cCnt++)
                    {
                        if (cCnt == 1)
                        {
                            u = (double)(testrange.Cells[rCnt + 1, 1] as Excel.Range).Value2;
                            Mainform1.Studentids.Add(u);
                        }
                        else if (cCnt == 2)
                        {
                            s = (string)(testrange.Cells[rCnt + 1, 2] as Excel.Range).Value2;
                            Mainform1.studentnames.Add(s);

                        }
                    }

                }
                startup1.Close(true, null, null);
                Startup1.Quit();

                Random random = new Random();
                int randomNumber = random.Next(14, 44);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Processing Data... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();


                Marshal.ReleaseComObject(Studentid1);
                Marshal.ReleaseComObject(startup1);
                Marshal.ReleaseComObject(Startup1);


                //NEXT WORKSHEET


                Excel.Worksheet ProdesCLO;


                input11 = Mainform1.input4 + "\\Marksheet Handling Program\\Setup File.xls";
                Startup1 = new Excel.Application();
                startup1 = Startup1.Workbooks.Open(@input11, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                ProdesCLO = (Excel.Worksheet)startup1.Worksheets.get_Item("Program Description");

                testrange = ProdesCLO.UsedRange;
                int lmn = 0, lmno = 0;
                while (lmn != -1)
                {
                    string courseadd = (string)(testrange.Cells[lmn + 1, 2] as Excel.Range).Value2;

                    if (courseadd == null)
                    {
                        lmn = -1;
                    }
                    else
                    {
                        lmn++;
                        lmno = lmn;
                    }
                };
                rw = lmno - 1;
                random = new Random();
                randomNumber = random.Next(44, 74);
                metroProgressBar1.Value = randomNumber;
                lblProcess.Text = "Processing Data... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();
                progress++;

                Startup1.Visible = false;
                string semes1 = "initialise";
                Mainform1.NumberOfSemesters = 0;
                for (rCnt = 1; rCnt <= rw; rCnt++)
                {

                    string semes = (string)(testrange.Cells[rCnt + 1, 1] as Excel.Range).Value2;

                    if (semes != null)
                    {
                        semes1 = semes;
                        Mainform1.NumberOfSemesters += 1;
                        Mainform1.semesterlist1.Add(new List<string>(Mainform1.NumberOfSemesters - 1));
                        Mainform1.semesterlist1[Mainform1.NumberOfSemesters - 1].Add(semes1);

                    }
                    string cours = (string)(testrange.Cells[rCnt + 1, 2] as Excel.Range).Value2;
                    int totalcount = 0;
                    for (int i = 0; i < Mainform1.NumberOfSemesters; i++)
                    {
                        string semest = "Semester " + Convert.ToString(i + 1);
                        totalcount += Mainform1.semesterlist1[i].Count - 1;
                        if (semes1 == semest && rCnt > totalcount)
                        {
                            Mainform1.semesterlist1[i].Add(cours);
                        }
                    }

                }
                random = new Random();
                randomNumber = random.Next(74, 94);
                metroProgressBar1.Value =randomNumber;
                lblProcess.Text = "Processing Data... " + randomNumber + " %";
                metroProgressBar1.Update();
                lblProcess.Update();
                progress++;
                for (rCnt = 1; rCnt <= rw; rCnt++)
                {
                    Mainform1.CLOs.Add(new List<string>(rCnt - 1));
                    for (cCnt = 2; cCnt <= 3; cCnt++)
                    {
                        if (cCnt == 2)
                        {
                            x = (string)(testrange.Cells[rCnt + 1, 2] as Excel.Range).Value2;
                            Mainform1.CLOs[rCnt - 1].Add(x);
                        }
                        else if (cCnt == 3)
                        {
                            double t = (double)(testrange.Cells[rCnt + 1, 3] as Excel.Range).Value2;
                            s = Convert.ToString(t);
                            Mainform1.CLOs[rCnt - 1].Add(s);
                        }
                    }

                }



                startup1.Close(true, null, null);
                Startup1.Quit();
                metroProgressBar1.Value = 100;
                lblProcess.Text = "Done! " + "100" + " %";
                metroProgressBar1.Update();
                lblProcess.Update();
                progress++;


                Marshal.ReleaseComObject(ProdesCLO);
                Marshal.ReleaseComObject(startup1);
                Marshal.ReleaseComObject(Startup1);
                #endregion
            }
            else if (Mainform1.outputcontrol==3)
            {
                #region Creating Seperate Marksheet Templates

                /*string pathstring = input4 + "\\Marksheet Handling Program";
                pathString2 = System.IO.Path.Combine(pathstring, "Marksheets");
                System.IO.Directory.CreateDirectory(pathString2);*/
                int klm = 0;
                progress = 0;
                int NoofCourses = 0;
                for (int j = 0; j < Mainform1.NumberOfSemesters; j++)
                {
                    NoofCourses = NoofCourses + (Mainform1.semesterlist1[j].Count - 1);


                }
                for (int j = 0; j < Mainform1.NumberOfSemesters; j++)
                {
                    Excel.Application Coursetemplate = new Excel.Application();
                    if (Coursetemplate == null)
                    {
                        MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                        return;
                    }

                    Coursetemplate.Visible = false; // To avoid error hresult 0x800ac472
                    Coursetemplate.DisplayAlerts = false;
                    Excel.Workbook coursetemplate = Coursetemplate.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                    for (int i = 1; i < Mainform1.semesterlist1[j].Count; i++)
                    {
                        Excel.Worksheet coursetemp = (Excel.Worksheet)coursetemplate.Worksheets.Add();
                        coursetemp.Name = Mainform1.semesterlist1[j][i] + " Marksheet";

                        Coursetemplate.Sheets[1].Range[Coursetemplate.Sheets[1].Cells[1, 1], Coursetemplate.Sheets[1].Cells[Mainform1.studentnames.Count + 3, int.Parse(Mainform1.CLOs[klm][1]) + 1]].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Coursetemplate.Sheets[1].Range[Coursetemplate.Sheets[1].Cells[1, 1], Coursetemplate.Sheets[1].Cells[Mainform1.studentnames.Count + 3, int.Parse(Mainform1.CLOs[klm][1]) + 1]].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                        for (int c = 1; c <= int.Parse(Mainform1.CLOs[klm][1]); c++)
                        {

                            coursetemp.Cells[1, 1].EntireColumn.columnwidth = 15;
                            coursetemp.Cells[1, 1].EntireRow.Font.Bold = true;
                            coursetemp.Cells[2, 1].Font.Bold = true;
                            coursetemp.Cells[3, 1].Font.Bold = true;
                            coursetemp.Cells[1, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                            coursetemp.Cells[1, c + 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                            coursetemp.Cells[2, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                            coursetemp.Cells[3, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                            coursetemp.Cells[1, 1].Value2 = "CLOs: ";
                            coursetemp.Cells[2, 1].Value2 = "Associated PLOs:";
                            coursetemp.Cells[3, 1].Value2 = "Student Names";
                            coursetemp.Cells[1, c + 1].Value2 = "CLO " + Convert.ToString(c);

                            for (int b = 1; b <= Mainform1.studentnames.Count; b++)
                            {
                                coursetemp.Cells[b + 3, c + 1].Value2 = "-";
                            }


                        }
                        for (int b = 1; b <= Mainform1.studentnames.Count; b++)
                        {
                            coursetemp.Cells[b + 3, 1].Value2 = Mainform1.studentnames[b - 1];
                        }
                        klm++;
                        Marshal.ReleaseComObject(coursetemp);
                        metroProgressBar1.Value = progress * 100 / NoofCourses;
                        lblProcess.Text = "Processing Data... " + (progress * 100 / NoofCourses) + " %";
                        metroProgressBar1.Update();
                        lblProcess.Update();
                        progress++;
                    }
                    string coursetempfilesave = Mainform1.pathString2 + "\\" + Mainform1.semesterlist1[j][0] + ".xls";
                    coursetemplate.SaveAs(@coursetempfilesave, Excel.XlFileFormat.xlWorkbookNormal, null, null, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);

                    coursetemplate.Close(true, null, null);
                    Coursetemplate.Quit();

                    Marshal.ReleaseComObject(coursetemplate);
                    Marshal.ReleaseComObject(Coursetemplate);

                }
                #endregion
            }
            else if(Mainform1.outputcontrol==4)
            {
                #region Read Marks from Excel Marksheets
                Mainform1.PLOWeightageCLO = new int[Mainform1.NumberOfSemesters, 12];
                Mainform1.PLOWeightageSem = new int[Mainform1.NumberOfSemesters, 12];
                int[] PLOWeightageCheck = new int[12];
                int klmno = 0;
                progress = 0;
                int NoofCourses=0;
                for (int j = 0; j < Mainform1.NumberOfSemesters; j++)
                {
                    NoofCourses = NoofCourses + (Mainform1.semesterlist1[j].Count - 1);


                }

                

                    for (int j = 0; j < Mainform1.NumberOfSemesters; j++)
                {

                    for (int a = 0; a < 12; a++)
                    {
                        Mainform1.PLOWeightageCLO[j, a] = 0;
                        Mainform1.PLOWeightageSem[j, a] = 0;
                        PLOWeightageCheck[a] = 0;

                    }

                    Mainform1.studentdata.Add(new List<List<List<List<double>>>>(j));
                    string input13 = Mainform1.pathString2 + "\\" + Mainform1.semesterlist1[j][0] + ".xls";
                    Excel.Application Coursetemplate = new Excel.Application();
                    if (Coursetemplate == null)
                    {
                        MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                        return;
                    }

                    Coursetemplate.Visible = false; // To avoid error hresult 0x800ac472
                    Coursetemplate.DisplayAlerts = false;
                    Excel.Workbook coursetemplate = Coursetemplate.Workbooks.Open(@input13, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                    for (int i = 1; i < Mainform1.semesterlist1[j].Count; i++)
                    {
                        Mainform1.studentdata[j].Add(new List<List<List<double>>>(i - 1));
                        Excel.Worksheet coursetemp = (Excel.Worksheet)coursetemplate.Worksheets.get_Item(Mainform1.semesterlist1[j][i] + " Marksheet");
                        Excel.Range courseRange = coursetemp.UsedRange;

                        Coursetemplate.Visible = false;
                        for (int b = 0; b < Mainform1.studentnames.Count; b++)
                        {
                            Mainform1.studentdata[j][i - 1].Add(new List<List<double>>(b));

                            int check = 0;
                            for (int d = 0; d < int.Parse(Mainform1.CLOs[klmno][1]); d++)
                            {
                                Coursetemplate.Visible = false;
                                string check1 = Convert.ToString(courseRange.Cells[b + 4, d + 2].Value2);
                                if (check1 == "-")
                                {
                                    check = 1;

                                }


                            }

                            for (int c = 0; c < int.Parse(Mainform1.CLOs[klmno][1]); c++)
                            {
                                Mainform1.studentdata[j][i - 1][b].Add(new List<double>(c));
                                if ((courseRange.Cells[2, c + 2] as Excel.Range).Value2 == null || (courseRange.Cells[2, c + 2] as Excel.Range).Value2 > 12 || (courseRange.Cells[2, c + 2] as Excel.Range).Value2 < 1)
                                {
                                    string PLOerror = null;
                                    if ((courseRange.Cells[2, c + 2] as Excel.Range).Value2 == null)
                                    {
                                        Coursetemplate.Visible = false;
                                        InvalidPLOCLO = c + 1;
                                        InvalidPLOcourse = Mainform1.semesterlist1[j][i];
                                        InvalidPLOsem = Mainform1.semesterlist1[j][0];
                                        InvalidPLOindex = 0;
                                        InvalidPLO plo2 = new InvalidPLO();
                                        plo2.ShowDialog();
                                        PLOerror = InvalidPLO.InvalidPLOcorrect;
                                       Coursetemplate.Visible = false;
                                        courseRange.Cells[2, c + 2].Value2 = Int32.Parse(PLOerror);
                                    }
                                    else if ((courseRange.Cells[2, c + 2] as Excel.Range).Value2 > 12 || (courseRange.Cells[2, c + 2] as Excel.Range).Value2 < 1)
                                    {
                                        InvalidPLOCLO = c + 1;
                                        InvalidPLOcourse = Mainform1.semesterlist1[j][i];
                                        InvalidPLOsem = Mainform1.semesterlist1[j][0];
                                        InvalidPLOindex = 1;
                                        Coursetemplate.Visible = false;
                                        InvalidPLO plo1 = new InvalidPLO();
                                        plo1.ShowDialog();
                                        PLOerror = InvalidPLO.InvalidPLOcorrect;
                                        Coursetemplate.Visible = false;
                                        courseRange.Cells[2, c + 2].Value2 = Int32.Parse(PLOerror);
                                    }
                                    else { }
                                    Coursetemplate.DisplayAlerts = false;
                                    coursetemplate.SaveCopyAs(@input13);
                                    Marshal.ReleaseComObject(coursetemp);
                                    coursetemplate.Close(false, null, null);
                                    Marshal.ReleaseComObject(coursetemplate);
                                    coursetemplate = Coursetemplate.Workbooks.Open(@input13, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                                    Coursetemplate.Visible = false;
                                    coursetemp = (Excel.Worksheet)coursetemplate.Worksheets.get_Item(Mainform1.semesterlist1[j][i] + " Marksheet");
                                    courseRange = coursetemp.UsedRange;
                                }

                                if ((courseRange.Cells[2, c + 2] as Excel.Range).Value2 == null || (courseRange.Cells[2, c + 2] as Excel.Range).Value2 > 12 || (courseRange.Cells[2, c + 2] as Excel.Range).Value2 < 1)
                                {
                                    Mainform1.CLOs[klmno][1] = Convert.ToString(int.Parse(Mainform1.CLOs[klmno][1]) - 1);
                                }
                                else
                                {
                                    if (b == 0)
                                    {
                                        Mainform1.PLOWeightageCLO[j, Convert.ToInt32(courseRange.Cells[2, c + 2].Value2 - 1)] += 1;
                                        PLOWeightageCheck[Convert.ToInt32(courseRange.Cells[2, c + 2].Value2 - 1)] = 1;
                                    }
                                    if (check == 1)
                                    {
                                        Mainform1.studentdata[j][i - 1][b][c].Add(0);
                                        Mainform1.studentdata[j][i - 1][b][c].Add(0);
                                    }
                                    else
                                    {
                                        Mainform1.studentdata[j][i - 1][b][c].Add((double)(courseRange.Cells[b + 4, c + 2] as Excel.Range).Value2);
                                        Mainform1.studentdata[j][i - 1][b][c].Add((double)(courseRange.Cells[2, c + 2] as Excel.Range).Value2);
                                    }
                                }
                            }


                        }

                        klmno++;
                        Marshal.ReleaseComObject(coursetemp);

                        for (int a = 0; a < 12; a++)
                        {
                            Mainform1.PLOWeightageSem[j, a] = Mainform1.PLOWeightageSem[j, a] + PLOWeightageCheck[a];
                            PLOWeightageCheck[a] = 0;
                        }
                    
                    metroProgressBar1.Value = progress * 100 /NoofCourses ;
                    lblProcess.Text = "Processing Data... " + (progress * 100 / NoofCourses) + " %";
                    metroProgressBar1.Update();
                    lblProcess.Update();
                    progress++;
                    }


                    coursetemplate.Close(true, null, null);
                    Coursetemplate.Quit();

                    Marshal.ReleaseComObject(coursetemplate);
                    Marshal.ReleaseComObject(Coursetemplate);
                  
                }
                #endregion
            }
            else if(Mainform1.outputcontrol==5)
            {
                #region Output To Excel
                // MessageBox.Show("Please wait while the results are being compiled.");
                progress = 0;
            Excel.Application OutputTemplate = new Excel.Application();

            int[] StudentsPassedinPLO = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] StudentsFailedinPLO = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            string[,] colormatch = new string[Mainform1.NumberOfSemesters, Mainform1.semesterlist1.Count - 1];
            OutputTemplate.Visible = false; // To avoid error hresult 0x800ac472
            OutputTemplate.DisplayAlerts = false;
            Excel.Workbook outputtemplate = OutputTemplate.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            for (int j = Mainform1.NumberOfSemesters - 1; j >= 0; j--)
            {

                Excel.Worksheet outputtemp = (Excel.Worksheet)outputtemplate.Worksheets.Add();
                outputtemp.Name = Mainform1.semesterlist1[j][0];
                OutputTemplate.Sheets[1].Range[OutputTemplate.Sheets[1].Cells[1, 1], OutputTemplate.Sheets[1].Cells[Mainform1.studentnames.Count * 12 + 130, Mainform1.semesterlist1.Count * 10]].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                OutputTemplate.Sheets[1].Range[OutputTemplate.Sheets[1].Cells[1, 1], OutputTemplate.Sheets[1].Cells[Mainform1.studentnames.Count * 12 + 130, Mainform1.semesterlist1.Count * 10]].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                outputtemp.Cells[1, Mainform1.semesterlist1[j].Count + 3].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[1, Mainform1.semesterlist1[j].Count + 3].Value2 = "PLO Status";
                outputtemp.Cells[1, Mainform1.semesterlist1[j].Count + 3].EntireColumn.Font.Bold = true;
                outputtemp.Cells[1, 1].EntireRow.Font.Bold = true;
                outputtemp.Cells[2, 1].Font.Bold = true;
                outputtemp.Cells[2, 2].Font.Bold = true;
                outputtemp.Cells[2, 3].Font.Bold = true;
                outputtemp.Cells[1, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[1, 3].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[2, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[2, 2].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[2, 3].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[1, 3].Value2 = "Courses: ";
                outputtemp.Range[outputtemp.Cells[1, 1], outputtemp.Cells[1, 2]].Merge();
                outputtemp.Cells[2, 3].Value2 = "PLOs";
                outputtemp.Cells[2, 2].Value2 = "Student Names";
                outputtemp.Cells[2, 1].Value2 = "Student IDs";
                outputtemp.Cells[1, 2].EntireColumn.columnwidth = 20;
                outputtemp.Cells[1, 3].EntireColumn.columnwidth = 20;
                outputtemp.Cells[1, 1].EntireColumn.columnwidth = 15;
                outputtemp.Cells[1, Mainform1.semesterlist1[j].Count + 3].EntireColumn.columnwidth = 20;
                outputtemp.Range[outputtemp.Cells[Mainform1.studentnames.Count * 12 + 3, 1], outputtemp.Cells[Mainform1.studentnames.Count * 12 + 3, 2]].Merge();
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 3, 1].Value2 = "Total Number of Students: " + Convert.ToString(Mainform1.studentnames.Count);
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 2].Value2 = "Students Passed";
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 3].Value2 = "Students Failed";
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 3, 1].EntireRow.Font.Bold = true;
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 1].EntireRow.Font.Bold = true;
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 3, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 2].Interior.Color = System.Drawing.Color.NavajoWhite;
                outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7, 3].Interior.Color = System.Drawing.Color.NavajoWhite;






                for (int i = 1; i < Mainform1.semesterlist1[j].Count; i++)
                {

                    outputtemp.Cells[1, i + 3].EntireColumn.columnwidth = 15;
                    outputtemp.Cells[1, i + 3].Interior.Color = System.Drawing.Color.NavajoWhite;
                    outputtemp.Cells[1, i + 3].Value2 = Mainform1.semesterlist1[j][i];




                    int h = 0;
                    for (int b = 2; b < Mainform1.studentnames.Count * 12 + 2; b = b + 12)
                    {
                        outputtemp.Cells[b + 1, 2].Value2 = Mainform1.studentnames[h];
                        outputtemp.Cells[b + 1, 1].Value2 = Mainform1.Studentids[h];
                        outputtemp.Range[outputtemp.Cells[b + 1, 1], outputtemp.Cells[b + 12, 1]].Merge();
                        outputtemp.Range[outputtemp.Cells[b + 1, 2], outputtemp.Cells[b + 12, 2]].Merge();


                        for (int a = 1; a <= 12; a++)
                        {

                            outputtemp.Cells[a + b, 3].Value2 = "PLO " + Convert.ToString(a);
                            if (Convert.ToString(StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseMarks / StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseLinks) != null || StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseMarks / StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseLinks != double.NaN)
                            {
                                outputtemp.Cells[a + b, i + 3].Value2 = Convert.ToString((StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseMarks / StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOTotalCourseLinks)) + " %";
                            }
                            else
                            {
                                outputtemp.Cells[a + b, i + 3].Value2 = "-";

                            }
                            if (Convert.ToString(StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterPasses / StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterLinks) != null || StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterPasses / StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterLinks != double.NaN)
                            {
                                outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Value2 = Convert.ToString((StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterPasses / StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOTotalSemesterLinks) * 100) + " %";
                            }
                            else
                            {
                                outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Value2 = "-";
                            }
                            if (StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOSemState == "Fail")
                            {
                                outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Interior.Color = System.Drawing.Color.Salmon;

                                StudentsFailedinPLO[a - 1]++;
                            }
                            else if (StudentManager.Students[h].Semesters[j].PLOsSemester[a - 1].PLOSemState == "Pass")
                            {
                                outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Interior.Color = System.Drawing.Color.PaleGreen;

                                StudentsPassedinPLO[a - 1]++;
                            }
                            else { outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Interior.Color = System.Drawing.Color.WhiteSmoke; outputtemp.Cells[a + b, Mainform1.semesterlist1[j].Count + 3].Value2 = "Untested"; }

                            if (StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOCourseState == "Fail")
                            {
                                outputtemp.Cells[a + b, i + 3].Interior.Color = System.Drawing.Color.Salmon;

                            }
                            else if (StudentManager.Students[h].Semesters[j].Courses[i - 1].PLOsCourses[a - 1].PLOCourseState == "Pass")
                            {
                                outputtemp.Cells[a + b, i + 3].Interior.Color = System.Drawing.Color.PaleGreen;
                            }

                            else { outputtemp.Cells[a + b, i + 3].Interior.Color = System.Drawing.Color.WhiteSmoke; outputtemp.Cells[a + b, i + 3].Value2 = "-"; }

                        }

                        h++;
                    }

                    for (int a = 1; a <= 12; a++)
                    {

                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7, 1].Value2 = "PLO " + Convert.ToString(a);
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7, 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7, 1].Font.Bold = true;
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7, 2].Value2 = StudentsPassedinPLO[a - 1];
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7, 3].Value2 = StudentsFailedinPLO[a - 1];
                        StudentsPassedinPLO[a - 1] = 0;
                        StudentsFailedinPLO[a - 1] = 0;


                    }
                }

                   

                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)outputtemp.ChartObjects(Type.Missing);
                Excel.ChartObject myChart;
                Excel.Chart chartPage;

                for (int ab = 1; ab <= 3; ab++)
                {
                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Interior.Color = System.Drawing.Color.NavajoWhite;
                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Font.Bold = true;
                    if (ab == 1)
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = "Average Course Marks";
                    }
                    else if (ab == 2)
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = "Variance";
                    }
                    else
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = "Standard Deviation";
                    }

                    for (int a = 1; a <= 12; a++)
                    {

                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 1].Value2 = "PLO " + Convert.ToString(a);
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 1].Font.Bold = true;
                        if (ab == 1)
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = StudentManager.PLOStatistics[a - 1].PLOSemesterStatistics[j].AverageCLOMarks;
                        }
                        else if (ab == 2)
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = StudentManager.PLOStatistics[a - 1].PLOSemesterStatistics[j].Variance;
                        }
                        else
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 7 + (ab * 14), 2].Value2 = StudentManager.PLOStatistics[a - 1].PLOSemesterStatistics[j].StandardDeviation;
                        }


                    }



                    myChart = xlCharts.Add(350, (Mainform1.studentnames.Count * 12 + (ab * 14) + 2) * 15 + 60, 550, 195);
                    chartPage = myChart.Chart;

                    chartRange = outputtemp.get_Range("A" + Convert.ToString(Mainform1.studentnames.Count * 12 + 6 + (ab * 14)), "B" + Convert.ToString(Mainform1.studentnames.Count * 12 + 6 + (ab * 14) + 13));
                    chartPage.SetSourceData(chartRange);
                    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                    chartPage.Legend.LegendEntries(1).LegendKey.Interior.Color = System.Drawing.Color.CornflowerBlue;
                    if (ab == 1)
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Average Course Marks (PLOs)", ValueTitle: "Average %age Course Marks");
                    }
                    else if (ab == 2)
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Variance (PLOs)", ValueTitle: "Variance");
                    }
                    else
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Standard Deviation (PLOs)", ValueTitle: "Standard Deviation");
                    }

                }
                    progress++;
                    metroProgressBar1.Value = progress * 100 / (3 * Mainform1.NumberOfSemesters);
                    lblProcess.Text = "Generating results... " + (progress * 100 / (3 * Mainform1.NumberOfSemesters)) + " %";
                    metroProgressBar1.Update();
                    lblProcess.Update();
                    int spacecount;
                if (Mainform1.semesterlist1[j].Count - 1 < 12)
                {
                    spacecount = 12;
                }
                else
                {
                    spacecount = Mainform1.semesterlist1[j].Count - 1;
                }
                for (int ab = 1; ab <= 3; ab++)
                {


                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Interior.Color = System.Drawing.Color.NavajoWhite;
                    outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Font.Bold = true;
                    if (ab == 1)
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = "Average Course Marks";
                    }
                    else if (ab == 2)
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = "Variance";
                    }
                    else
                    {
                        outputtemp.Cells[Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = "Standard Deviation";
                    }

                    for (int a = 1; a < Mainform1.semesterlist1[j].Count; a++)
                    {

                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 1].Value2 = Mainform1.semesterlist1[j][a];
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 1].Interior.Color = System.Drawing.Color.NavajoWhite;
                        outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 1].Font.Bold = true;
                        if (ab == 1)
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = Math.Round(StudentManager.SemesterCourseStatistics[j].CourseAverages[a - 1],5);
                        }
                        else if (ab == 2)
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = Math.Round(StudentManager.SemesterCourseStatistics[j].CourseVariances[a - 1],5);
                        }
                        else
                        {
                            outputtemp.Cells[a + Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount)), 2].Value2 = Math.Round(StudentManager.SemesterCourseStatistics[j].CourseStandardDeviations[a - 1],5);
                        }


                    }


                    myChart = xlCharts.Add(350, (Mainform1.studentnames.Count * 12 + (ab * (2 + spacecount)) + 46) * 15 + 60, (1 + spacecount) * 46, 195);
                    chartPage = myChart.Chart;

                    chartRange = outputtemp.get_Range("A" + Convert.ToString(Mainform1.studentnames.Count * 12 + 51 + (ab * (2 + spacecount))), "B" + Convert.ToString(Mainform1.studentnames.Count * 12 + 50 + (ab * (2 + spacecount)) + Mainform1.semesterlist1[j].Count));
                    chartPage.SetSourceData(chartRange);
                    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                    chartPage.Legend.LegendEntries(1).LegendKey.Interior.Color = System.Drawing.Color.Tomato;

                    if (ab == 1)
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Average Course Marks (Courses)", ValueTitle: "Average %age Course Marks");
                    }
                    else if (ab == 2)
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Variance (Courses)", ValueTitle: "Variance");
                    }
                    else
                    {
                        chartPage.ChartWizard(Source: chartRange, Title: "Standard Deviation (Courses)", ValueTitle: "Standard Deviation");
                    }


                }

                int position = Mainform1.studentnames.Count * 12 + 74 + (3 * spacecount);

                for (int a = Mainform1.studentnames.Count; a >= 1; a--)
                {

                    outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, 1].Value2 = Mainform1.studentnames[a - 1];
                    outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, 1].Font.Color = System.Drawing.Color.White;

                    for (int b = Mainform1.semesterlist1[j].Count - 1; b >= 1; b--)
                    {
                        outputtemp.Cells[position, Mainform1.semesterlist1[j].Count - b + 1].Value2 = Mainform1.semesterlist1[j][b];
                        outputtemp.Cells[position, Mainform1.semesterlist1[j].Count - b + 1].Font.Color = System.Drawing.Color.White;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, Mainform1.semesterlist1[j].Count - b + 1].Value2 = Math.Round(StudentManager.Students[a - 1].Semesters[j].Courses[b - 1].CourseAverage,5);
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, Mainform1.semesterlist1[j].Count - b + 1].Font.Color = System.Drawing.Color.White;
                        outputtemp.Cells[position, (2 * Mainform1.semesterlist1[j].Count) - b].Value2 = Mainform1.semesterlist1[j][b];
                        outputtemp.Cells[position, (2 * Mainform1.semesterlist1[j].Count) - b].Font.Color = System.Drawing.Color.White;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count) - b].Value2 = Math.Round(StudentManager.SemesterCourseStatistics[j].CourseAverages[b - 1],5);
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count) - b].Font.Color = System.Drawing.Color.White;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count)].Value2 = 0;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count)].Font.Color = System.Drawing.Color.White;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count) + 1].Value2 = Mainform1.studentnames.Count;
                        outputtemp.Cells[Mainform1.studentnames.Count - a + 1 + position, (2 * Mainform1.semesterlist1[j].Count) + 1].Font.Color = System.Drawing.Color.White;
                    }


                }
                    progress++;
                    metroProgressBar1.Value = progress * 100 / (3 * Mainform1.NumberOfSemesters);
                    lblProcess.Text = "Generating results... " + (progress * 100 / (3 * Mainform1.NumberOfSemesters)) + " %";
                    metroProgressBar1.Update();
                    lblProcess.Update();

                myChart = xlCharts.Add((Mainform1.semesterlist1[j].Count) * 83 + 330, 17, ((Mainform1.studentnames.Count) * 12) * 15 + 12 + 47, 550);
                chartPage = myChart.Chart;
                chartPage.ChartType = Excel.XlChartType.xlBarClustered;
                chartPage.HasTitle = false;

                var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                yAxis.HasTitle = true;
                yAxis.AxisTitle.Text = "Average Marks Obtained (%)";
                yAxis.AxisTitle.Orientation = Excel.XlOrientation.xlHorizontal;
                yAxis.HasMinorGridlines = true;

                var xAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                xAxis.HasTitle = false;
                xAxis.CategoryNames = false;
                xAxis.HasMajorGridlines = true;
                Excel.SeriesCollection oSeriesCollection = (Excel.SeriesCollection)chartPage.SeriesCollection();
                Excel.Range xValRange = outputtemp.Range[outputtemp.Cells[position + 1, 1], outputtemp.Cells[position + Mainform1.studentnames.Count, 1]];
                chartPage.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowNone, false, false, true, false, false, false, false);

                for (int i = 0; i < Mainform1.semesterlist1[j].Count - 1; i++)
                {
                    Excel.Series oSeries = oSeriesCollection.NewSeries();
                    oSeries.Values = outputtemp.Range[outputtemp.Cells[position + 1, 2 + i], outputtemp.Cells[position + Mainform1.studentnames.Count, 2 + i]];
                    oSeries.XValues = xValRange;
                    oSeries.Name = Mainform1.semesterlist1[j][Mainform1.semesterlist1[j].Count - i - 1];
                    oSeries.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeMinusValues, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, StudentManager.SemesterCourseStatistics[j].CourseAverages[i], StudentManager.SemesterCourseStatistics[j].CourseAverages[i]);


                }




                for (int a = 1; a <= 12; a++)
                {

                    outputtemp.Cells[1 + a, Mainform1.semesterlist1[j].Count + 5].Value2 = "PLO " + Convert.ToString(a);
                    outputtemp.Cells[1 + a, Mainform1.semesterlist1[j].Count + 5].Font.Color = System.Drawing.Color.White;
                    outputtemp.Cells[1 + a, Mainform1.semesterlist1[j].Count + 6].Value2 = Mainform1.PLOWeightageCLO[j, a - 1];
                    outputtemp.Cells[1 + a, Mainform1.semesterlist1[j].Count + 6].Font.Color = System.Drawing.Color.White;
                    outputtemp.Cells[1 + a + 12, Mainform1.semesterlist1[j].Count + 5].Value2 = "PLO " + Convert.ToString(a);
                    outputtemp.Cells[1 + a + 12, Mainform1.semesterlist1[j].Count + 5].Font.Color = System.Drawing.Color.White;
                    outputtemp.Cells[1 + a + 12, Mainform1.semesterlist1[j].Count + 6].Value2 = Mainform1.PLOWeightageSem[j, a - 1];
                    outputtemp.Cells[1 + a + 12, Mainform1.semesterlist1[j].Count + 6].Font.Color = System.Drawing.Color.White;

                }
                myChart = xlCharts.Add(950, (Mainform1.studentnames.Count * 12) * 15 + 90, 400, 360);
                chartPage = myChart.Chart;
                chartRange = outputtemp.Range[outputtemp.Cells[2, Mainform1.semesterlist1[j].Count + 5], outputtemp.Cells[13, Mainform1.semesterlist1[j].Count + 6]];
                chartPage.SetSourceData(chartRange);
                chartPage.ChartType = Excel.XlChartType.xlDoughnut;
                Excel.Axis axis = chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary) as Excel.Axis;
                chartPage.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent, true, true, true, false, false, false, true);
                chartPage.ChartWizard(Source: chartRange, Title: "PLO Weightage (w.r.t) CLOs");

                myChart = xlCharts.Add(950, (Mainform1.studentnames.Count * 12) * 15 + 480, 400, 350);
                chartPage = myChart.Chart;
                chartRange = outputtemp.Range[outputtemp.Cells[14, Mainform1.semesterlist1[j].Count + 5], outputtemp.Cells[25, Mainform1.semesterlist1[j].Count + 6]];
                chartPage.SetSourceData(chartRange);
                chartPage.ChartType = Excel.XlChartType.xlDoughnut;
                axis = chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary) as Excel.Axis;
                chartPage.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent, true, true, true, false, false, false, true);
                chartPage.ChartWizard(Source: chartRange, Title: "PLO Weightage (w.r.t) Courses");

                myChart = xlCharts.Add(350, (Mainform1.studentnames.Count * 12) * 15 + 90, 550, 195);
                chartPage = myChart.Chart;
                chartRange = outputtemp.get_Range("A" + Convert.ToString(Mainform1.studentnames.Count * 12 + 7), "C" + Convert.ToString(Mainform1.studentnames.Count * 12 + 19));

                chartPage.SetSourceData(chartRange);
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.ChartWizard(Source: chartRange, Title: "Pass/Fail Statistics", ValueTitle: "Number of Students");
                chartPage.Legend.LegendEntries(1).LegendKey.Interior.Color = System.Drawing.Color.PaleGreen;
                chartPage.Legend.LegendEntries(2).LegendKey.Interior.Color = System.Drawing.Color.Salmon;


                Marshal.ReleaseComObject(outputtemp);
                progress++;
                metroProgressBar1.Value = progress * 100 /(3* Mainform1.NumberOfSemesters);
                lblProcess.Text = "Generating results... " + (progress * 100 /(3* Mainform1.NumberOfSemesters)) + " %";
                metroProgressBar1.Update();
                lblProcess.Update();
               
            }
            string outputtempfilesave = Mainform1.pathString2 + "\\Results.xls";
            outputtemplate.SaveAs(outputtempfilesave, Excel.XlFileFormat.xlWorkbookNormal, null, null, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);
            outputtemplate.Close(true, null, null);
            OutputTemplate.Quit();


            Marshal.ReleaseComObject(outputtemplate);
            Marshal.ReleaseComObject(OutputTemplate);

            System.Diagnostics.Process.Start("explorer.exe", Mainform1.pathString2);
            #endregion
            }
            else { }
            

            Close();


        }

        private void ProgressBar_Load(object sender, EventArgs e)
        {

        }
    }
}
