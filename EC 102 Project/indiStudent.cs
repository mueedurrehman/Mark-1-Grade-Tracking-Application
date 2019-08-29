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
    public partial class indiStudent : MetroFramework.Forms.MetroForm
    {
        public indiStudent()
        {
            InitializeComponent();
        }

        private void indiStudent_Load(object sender, EventArgs e)
        {
            metroComboBox1.Items.Clear();
            for(int j=0;j<Mainform1.semesterlist1.Count;j++)
            {
                metroComboBox1.Items.Add(Convert.ToString(j + 1));
            }
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {

            timer1.Start();
        }
        int studentmatch;
        private void timer1_Tick(object sender, EventArgs e)
        {
           
                    if (Int32.TryParse(metroTextBox1.Text,out studentmatch))
                    {
                         for (int stIndex = 0; stIndex < StudentManager.Students.Count; stIndex++)
                        {

                            if (StudentManager.Students[stIndex].StudentID == studentmatch)
                            {
                                studentmatch = stIndex;
                                pictureBox1.Visible = false;
                                metroLabel3.Visible = false;
                                metroButton1.Enabled = true;
                                break;
                            }
                            else
                            {
                                 pictureBox1.Visible = true;
                                 metroLabel3.Visible = true;
                                 metroButton1.Enabled = false;
                            }

                        }
                     
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(metroTextBox1.Text))
                {
                            pictureBox1.Visible = false;
                            metroLabel3.Visible = false;
                        }
                         else
                        {
                            pictureBox1.Visible = true;
                            metroLabel3.Visible = true;
                        }
                       
                        metroButton1.Enabled = false;
                    }

                
            
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            metroLabel5.Visible = true;
            metroPanel1.Visible = true;
            var student = StudentManager.Students[studentmatch];
            string text2="Student Name: " + StudentManager.Students[studentmatch].Name;
                 var semester = student.Semesters[Convert.ToInt32(metroComboBox1.Text) - 1];
            text2=text2+"\nSemester Number: " + semester.SemNum;
            for (int SemPLOIndex = 0; SemPLOIndex < 12; SemPLOIndex++)
            {
                var SemPLOStat = semester.PLOsSemester[SemPLOIndex];
                if (SemPLOStat.PLOTotalSemesterLinks != 0)
                {
                    text2=text2+"\nPLO " +Convert.ToString(SemPLOIndex + 1 )+ ": " + Convert.ToString(SemPLOStat.PLOSemState);
                }
            }
            for (int cIndex = 0; cIndex < semester.TotalCourses; cIndex++)
            {
                var course = semester.Courses[cIndex];
                text2=text2 +"\n\nCourse ID: " + course.CourseID;
                for (int cloIndex = 0; cloIndex < course.CLOs.Count; cloIndex++)
                {
                    var clo = course.CLOs[cloIndex];
                   
                    if (course.CLOs[cloIndex].PLOPass == 1)
                    {
                        text2=text2+"\nCLO " + Convert.ToString(course.CLOs[cloIndex].CLOnumber) +": " + Convert.ToString(clo.CLOmarks) +"  Passed" ;
                    }
                    else
                    {
                        text2 = text2 + "\nCLO " + Convert.ToString(course.CLOs[cloIndex].CLOnumber) + ": " + Convert.ToString(clo.CLOmarks) + "  Failed"; ;

                    }
                }
                text2 = text2 + "\n";
                for (int CoursePLOIndex = 0; CoursePLOIndex < 12; CoursePLOIndex++ )
                {
                    var CoursePLOStat = course.PLOsCourses[CoursePLOIndex];
                    if (CoursePLOStat.PLOTotalCourseLinks != 0)
                    {
                        text2=text2+"\nPLO " + Convert.ToString(CoursePLOIndex+1) + " State for Course "+ course.CourseID + " : "+ CoursePLOStat.PLOCourseState; // Should use placeholder and then add 1.
                    }
                }
            }

            metroLabel4.Text = text2;
        }

    }
}
