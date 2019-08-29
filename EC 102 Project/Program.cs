using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EC_102_Project;
using System.Windows.Forms;

//declaration of classes that create the overall storage framework for this project
public class SemesterStatistic
{
    public List<double> CourseAverages { get; set; }
    public List<double> CourseStandardDeviations { get; set; }
    public List<double> CourseVariances { get; set; }
}
public class PLOStatistic
{
    public List<PLOSemesterStatistic> PLOSemesterStatistics { get; set; }
}
public class PLOSemesterStatistic
{
    public PLOSemesterStatistic()
    {
        TotalCLOMarks = 0;
        SumofDeviationsSquared = 0;
    }
    public double TotalCLOMarks { get; set; }
    public double TotalCLOs { get; set; }
    public List<int> CLOMarks { get; set; }
    public double SumofDeviationsSquared { get; set; }
    public double AverageCLOMarks { get; set; }
    public double StandardDeviation { get; set; }
    public double Variance { get; set; }
}
public class Student
{
    public string Name { get; set; }
    public int StudentID { get; set; }
    public List<Semester> Semesters { get; set; }
}

public class Semester
{
    public int SemNum { get; set; }
    public int TotalCourses { get; set; }
    public List<Course> Courses { get; set; }
    public List<PLO> PLOsSemester { get; set; }
}

public class Course
{
    public Course()
    {
        TotalCourseMarks = 0;
        TotalCLOs = 0;
    }
    public string CourseID { get; set; }
    public string CourseName { get; set; }
    public int TotalCourseMarks { get; set; }
    public double CourseAverage { get; set; }
    public double TotalCLOs { get; set; }
    public List<CLO> CLOs { get; set; }
    public List<PLO> PLOsCourses { get; set; }
}

public class CLO
{
    public int PLOLink { get; set; }
    public int PLOPass { get; set; }
    public int CLOnumber { get; set; }
    //public string CLOdescription { get; set; }
    public int CLOmarks { get; set; }
}

public class PLO
{
    public PLO()
    {
        PLOTotalSemesterPasses = 0;
        PLOTotalSemesterLinks = 0;
        PLOTotalCourseLinks = 0;
        PLOTotalCoursePasses = 0;
    }
    public int PLONumber { get; set; }
    public double PLOTotalSemesterLinks { get; set; }
    public double PLOTotalSemesterPasses { get; set; }
    public string PLOSemState { get; set; }
    public double PLOTotalCourseLinks { get; set; }
    public double PLOTotalCourseMarks { get; set; }
    public List<int> CLOMarksforDeviations = new List<int>();
    public double PLOTotalCoursePasses { get; set; }
    public string PLOCourseState { get; set; }

}

//the class StudentManager houses the methods that will be used for program initiation and execution
public class StudentManager
{
    //A list of type Student that will be used to store all pertinent data on every student
    public static List<Student> Students = new List<Student>();
    public static List<PLOStatistic> PLOStatistics = new List<PLOStatistic>();
    public static List<SemesterStatistic> SemesterCourseStatistics = new List<SemesterStatistic>();
    //This method, when called, adds an object of type student to the list declared above.
    public void AddNewStudent(Student student)
    {
        Students.Add(student);
    }

    public void Start()
    {
        //Presentation of input options to the user
        int UserOption = 0;
        Console.WriteLine("Welcome to the Marksheet Handling Program");
        Console.WriteLine("Do you wish to enter data using console or read data from an excel file?");
        Console.WriteLine("Press 1 to enter data via the console, 2 to enter data via an excel file");
        Console.WriteLine("If you wish to output processed results to an excel file and/or wish to create graphs, data must be input via an excel file.");

        /*The structure below has been used multiple times throughout the Program.cs file. Essentially, it is meant to ensure that
        integer inputs are of the correct type and fall within the expected range so program does not receive erroneous data. Error
        messages are generated repeatedly if need be.*/
        
        UserOption = 3;
       
        while (UserOption != -1)
        {
            switch (UserOption)
            {
                case 3:
                    Welcome welcome = new Welcome();
                    welcome.ShowDialog();
                    Mainform1 f1 = new Mainform1();
                    f1.ShowDialog();      //Passes control over to the Windows Forms portion to allow for excel file creation.
                    //MessageBox.Show(Convert.ToString(OutputExcel.outputagain) + " " + Convert.ToString(OutputExcel.outputagain2));
                    ConstructStudentExcel();        //like ConstructStudent, this stores the Excel Data in the appropriate arrays.
                    PLOStatisticCalculations();
                    SemesterCourseAverageCalculation();
                    UserOption = MenuNavigation();      //Another menu navigation method that is excel specific in terms of allowed options.
                    while ( UserOption!=-1)
                    {
                        switch (UserOption)
                        {
                            case 1:
                                Mainform1.outputcontrol = 11;
                                OutputExcel f4 = new OutputExcel(); //Passing control to Form2 object. This allows for excel output and graph creation.
                                f4.ShowDialog();//This method prints data for specific students only
                                break;
                           
                            case 3:
                                Mainform1.outputcontrol = 5;
                                OutputExcel f2 = new OutputExcel(); //Passing control to Form2 object. This allows for excel output and graph creation.
                                f2.ShowDialog();
                                break;
                           
                            case 5:
                                PrintAllPLOAllSemesterStatistics();
                                break;
                            case 6:
                                PrintAllPLOSpecificSemesterStatistics();
                                break;
                            case 7:
                                PrintSpecificPLOAllSemesterStatistics();
                                break;
                            case 8:
                                Mainform1.outputcontrol = 10;
                                OutputExcel f3 = new OutputExcel(); //Passing control to Form2 object. This allows for excel output and graph creation.
                                f3.ShowDialog();
                                break;
                        }

                        UserOption = MenuNavigation();

                    }
                    break;
            }
            Students.Clear(); //Clearing the Students list of data references.
            PLOStatistics.Clear();
            GC.Collect(); // Manual garbage collection to remove data that has no references.
            Console.WriteLine("Do you wish to enter data using console or read data from an excel file?");
            Console.WriteLine("Press 1 to enter data via the console, 2 to enter data via an excel file");
            Console.WriteLine("Data read from an excel file will be output to an excel file that will contain the graphs as well");
            Console.WriteLine("If your work is done, enter -1 to exit the program");
            
            UserOption = OutputExcel.outputagain2;

          
        }
        Environment.Exit(0);
    } //NEED TO ADD OTHER PRINTING FUNCTIONS.
    public void SemesterCourseAverageCalculation()
    {
        for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
        {
            SemesterStatistic SemesterCourseStat = new SemesterStatistic();
            SemesterCourseStat.CourseAverages = new List<double>();
            SemesterCourseStat.CourseVariances = new List<double>();
            SemesterCourseStat.CourseStandardDeviations = new List<double>();
            int courseIndex = 0;
            while (courseIndex < Students[0].Semesters[SemesterIndex].Courses.Count)
            {
                double runningAverageTotal = 0;
                double runningCourseAttemptsTotal = 0;
                double runningAverageFinal = 0;
                double SumOfDeviationsSquared = 0;
                double CourseStandardDeviation = 0;
                double CourseVariance = 0;
                List<double> CourseAverageForDeviations = new List<double>();
                for (int StudentIndex = 0; StudentIndex < Students.Count; StudentIndex++)
                {
                    if (Students[StudentIndex].Semesters[SemesterIndex].Courses[courseIndex].CourseAverage != 0)
                    {
                        CourseAverageForDeviations.Add(Students[StudentIndex].Semesters[SemesterIndex].Courses[courseIndex].CourseAverage);
                        runningAverageTotal += Students[StudentIndex].Semesters[SemesterIndex].Courses[courseIndex].CourseAverage;
                        runningCourseAttemptsTotal += 1;
                    }
                }
                runningAverageFinal = runningAverageTotal / runningCourseAttemptsTotal;
                SemesterCourseStat.CourseAverages.Add(runningAverageFinal);
                for (int DeviationsIndex = 0; DeviationsIndex < CourseAverageForDeviations.Count; DeviationsIndex++)
                {
                    SumOfDeviationsSquared += SumOfDeviationsSquared + ((CourseAverageForDeviations[DeviationsIndex] - runningAverageFinal) * (CourseAverageForDeviations[DeviationsIndex] - runningAverageFinal));
                }
                if (SumOfDeviationsSquared != 0)
                {
                    CourseVariance = SumOfDeviationsSquared / (runningCourseAttemptsTotal - 1);
                    CourseStandardDeviation = Math.Sqrt(CourseVariance);
                }
                SemesterCourseStat.CourseVariances.Add(Math.Round(CourseVariance,2));
                SemesterCourseStat.CourseStandardDeviations.Add(Math.Round(CourseStandardDeviation,2));
                courseIndex++;
            }
            SemesterCourseStatistics.Add(SemesterCourseStat);
        }
    }
    public void PrintAllPLOAllSemesterStatistics()
    {
        for (int PLOIndex = 0; PLOIndex < 12; PLOIndex++)
        {
            for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
            {
                Console.WriteLine("PLO {0} Statistics for Semester {1}", PLOIndex + 1, SemesterIndex + 1);
                Console.WriteLine();
                Console.WriteLine("For PLO {0}, in Semester {1}, the Average CLO Marks are {2}", PLOIndex + 1, SemesterIndex + 1, PLOStatistics[PLOIndex].PLOSemesterStatistics[SemesterIndex].AverageCLOMarks);
                Console.WriteLine("For PLO {0}, in Semester {1}, Variance in Average CLO Marks is {2}", PLOIndex + 1, SemesterIndex + 1, PLOStatistics[PLOIndex].PLOSemesterStatistics[SemesterIndex].Variance);
                Console.WriteLine("For PLO {0}, in Semester {1}, the Standard Deviation in Average CLO Marks is {2}", PLOIndex + 1, SemesterIndex + 1, PLOStatistics[PLOIndex].PLOSemesterStatistics[SemesterIndex].StandardDeviation);
                Console.WriteLine();
            }
        }
    } // Need to Test AND MAKE MORE MODULAR. Specific Semesters and specific PLOs can be made common.
    public void PrintSpecificPLOSpecificSemesterStatistics() //Need to Test
    {
        //int InputChecker = 0;
        //bool PrintMorePLOs = true;
        ////bool PrintMoreSemesters = true;
        //bool falseIntInput;
        //string UserResponse;
        List<int> SelectedPLOs = new List<int>();
        List<int> SelectedSem = new List<int>();
        //Console.WriteLine("For which PLOs do you wish to output data. Please enter the PLO's number: ");
        //while (PrintMorePLOs)
        //{
        //    falseIntInput = true;
        //    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
        //    {
        //        if (1 <= InputChecker || InputChecker <= 12)
        //        {
        //            falseIntInput = false;
        //            break;
        //        }
        //        Console.WriteLine("Incorrect Input.Please enter a number from 1 to 12");
        //    }
        //SelectedPLOs.Add(InputChecker);
        //Console.Write("Do you wish to add another PLO.");
        //UserResponse = Console.ReadLine();
        //    while (UserResponse != "y" && UserResponse != "n")
        //    {
        //    Console.WriteLine("Incorrect Input. Please enter y to add another PLO or n to proceed further: ");
        //    UserResponse = Console.ReadLine();
        //    }
        //    if (UserResponse == "n")
        //    {
        //    PrintMorePLOs = false;
        //    }
        //}
        //SelectedPLOs = SelectedPLOs.OrderBy(x => x).ToList();
        //Console.Write("For which semester(s) do you wish to output data. Please enter the semester's number: ");
        //while (PrintMoreSemesters)
        //{
        //    falseIntInput = true;
        //    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
        //    {
        //        if (1 <= InputChecker || InputChecker <= 8)
        //        {
        //            falseIntInput = false;
        //            break;
        //        }
        //        Console.WriteLine("Incorrect Input.Please enter a number from 1 to 8");
        //    }
        //    SelectedSem.Add(InputChecker);
        //    Console.Write("Do you wish to add another semester: (y/n) ");
        //    UserResponse = Console.ReadLine();
        //    while (UserResponse != "y" && UserResponse != "n")
        //    {
        //        Console.WriteLine("Incorrect Input. Please enter y to select another semester or n to proceed further: ");
        //        UserResponse = Console.ReadLine();
        //    }
        //    if (UserResponse == "n")
        //    {
        //        PrintMoreSemesters = false;
        //    }
        //}
        //SelectedSem = SelectedSem.OrderBy(x => x).ToList();
        
        SelectedPLOs.AddRange(PLOstatistics.PLOnumform);
        SelectedSem.AddRange(PLOstatistics.semesnumform);
        MessageBox.Show("Hello");
        for (int PLOIndex = 0; PLOIndex < SelectedPLOs.Count; PLOIndex++)
        {
            for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
            {
                int actualSem = SelectedSem[semIndex] - 1;
                int actualPLO = SelectedPLOs[PLOIndex] - 1;
                var statistic = PLOStatistics[actualPLO].PLOSemesterStatistics[actualSem];
                for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
                {
                    Console.WriteLine("PLO {0} Statistics for Semester {1}", PLOIndex + 1, SelectedSem[semIndex]);
                    Console.WriteLine();
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Average CLO Marks are {2}", actualPLO + 1, SelectedSem[actualSem], statistic.AverageCLOMarks);
                    Console.WriteLine("For PLO {0}, in Semester {1}, Variance in Average CLO Marks is {2}", actualPLO + 1, SelectedSem[actualSem], statistic.Variance);
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Standard Deviation in Average CLO Marks is {2}", actualPLO + 1, SelectedSem[actualSem], statistic.StandardDeviation);
                    Console.WriteLine();
                }
            }
        }
    }
    public void PrintAllPLOSpecificSemesterStatistics()
    {
        //int InputChecker = 0;
        //bool PrintMoreSemesters = true;
        //bool falseIntInput;
        //string UserResponse;
        List<int> SelectedSem = new List<int>();
        //Console.Write("For which semester(s) do you wish to output data. Please enter the semester's number: ");
        //while (PrintMoreSemesters)
        //{
        //    falseIntInput = true;
        //    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
        //    {
        //        if (1 <= InputChecker || InputChecker <= 8)
        //        {
        //            falseIntInput = false;
        //            break;
        //        }
        //        Console.WriteLine("Incorrect Input.Please enter a number from 1 to 8");
        //    }
        //    SelectedSem.Add(InputChecker);
        //    Console.Write("Do you wish to add another semester: (y/n) ");
        //    UserResponse = Console.ReadLine();
        //    while (UserResponse != "y" && UserResponse != "n")
        //    {
        //        Console.WriteLine("Incorrect Input. Please enter y to select another semester or n to proceed further: ");
        //        UserResponse = Console.ReadLine();
        //    }
        //    if (UserResponse == "n")
        //    {
        //        PrintMoreSemesters = false;
        //    }
        //}
        //SelectedSem = SelectedSem.OrderBy(x => x).ToList();
        SelectedSem = SelectSpecificSemester();
        for (int PLOIndex = 0; PLOIndex < 12; PLOIndex++)
        {
            for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
            {
                var statistic = PLOStatistics[PLOIndex].PLOSemesterStatistics[SelectedSem[semIndex] - 1];
                for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
                {
                    Console.WriteLine("PLO {0} Statistics for Semester {1}", PLOIndex + 1, SelectedSem[semIndex]);
                    Console.WriteLine();
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Average CLO Marks are {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.AverageCLOMarks);
                    Console.WriteLine("For PLO {0}, in Semester {1}, Variance in Average CLO Marks is {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.Variance);
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Standard Deviation in Average CLO Marks is {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.StandardDeviation);
                    Console.WriteLine();
                }
            }
        }
    } // Need to Test
    public void PrintSpecificPLOAllSemesterStatistics()
    {
        //int InputChecker = 0;
        //bool PrintMorePLOs = true;
        //bool falseIntInput;
        //string UserResponse;
        List<int> SelectedPLOs = new List<int>();
        //Console.WriteLine("For which PLOs do you wish to output data. Please enter the PLO's number: ");
        //while (PrintMorePLOs)
        //{
        //    falseIntInput = true;
        //    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
        //    {
        //        if (1 <= InputChecker || InputChecker <= 12)
        //        {
        //            falseIntInput = false;
        //            break;
        //        }
        //        Console.WriteLine("Incorrect Input.Please enter a number from 1 to 12");
        //    }
        //    SelectedPLOs.Add(InputChecker);
        //    Console.Write("Do you wish to add another PLO.");
        //    UserResponse = Console.ReadLine();
        //    while (UserResponse != "y" && UserResponse != "n")
        //    {
        //        Console.WriteLine("Incorrect Input. Please enter y to add another PLO or n to proceed further: ");
        //        UserResponse = Console.ReadLine();
        //    }
        //    if (UserResponse == "n")
        //    {
        //        PrintMorePLOs = false;
        //    }
        //}
        //SelectedPLOs = SelectedPLOs.OrderBy(x => x).ToList();
        SelectedPLOs.AddRange(SelectSpecificPLOs());
        for (int PLOIndex = 0; PLOIndex < SelectedPLOs.Count; PLOIndex++)
        {
            for (int semIndex = 0; semIndex < 8; semIndex++)
            {
                int actualPLO = SelectedPLOs[PLOIndex] - 1;
                var statistic = PLOStatistics[actualPLO].PLOSemesterStatistics[semIndex];
                for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
                {
                    Console.WriteLine("PLO {0} Statistics for Semester {1}", PLOIndex + 1, SemesterIndex + 1);
                    Console.WriteLine();
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Average CLO Marks are {2}", actualPLO + 1, semIndex + 1, statistic.AverageCLOMarks);
                    Console.WriteLine("For PLO {0}, in Semester {1}, Variance in Average CLO Marks is {2}", actualPLO + 1, semIndex + 1, statistic.Variance);
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Standard Deviation in Average CLO Marks is {2}", actualPLO + 1, semIndex + 1, statistic.StandardDeviation);
                    Console.WriteLine();
                }
            }
        }
    } // Need to Test
    public void PrintAllSemesterCourseStatistics()
    {
        for (int SemesterIndex = 0; SemesterIndex < Students[0].Semesters.Count; SemesterIndex++)
        {
            Console.WriteLine("Course Statistics for Semester {0}", SemesterIndex + 1);
            Console.WriteLine();
            for (int CourseIndex = 0; CourseIndex < Students[0].Semesters[SemesterIndex].Courses.Count; CourseIndex++)
            {
                Console.WriteLine("Semester {0} Course {1}, Course Average: {2}", SemesterIndex + 1, CourseIndex + 1, SemesterCourseStatistics[SemesterIndex].CourseAverages[CourseIndex]);
                Console.WriteLine("Semester {0} Course {1}, Course Variance: {2}", SemesterIndex + 1, CourseIndex + 1, SemesterCourseStatistics[SemesterIndex].CourseVariances[CourseIndex]);
                Console.WriteLine("Semester {0} Course {1}, Course Standard Deviation: {2}", SemesterIndex + 1, CourseIndex + 1, SemesterCourseStatistics[SemesterIndex].CourseStandardDeviations[CourseIndex]);
                Console.WriteLine();
            }
        }
    }
    public void PrintSpecificSemesterCourseStatistics()
    {
        List<int> SelectedSem = new List<int>();
        SelectedSem.AddRange(SelectSpecificSemester());
        for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
        {
            for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
            {
                int currentSem = SelectedSem[SemesterIndex] - 1;
                Console.WriteLine("Course Statistics for Semester {0}", SemesterIndex + 1);
                Console.WriteLine();
                for (int CourseIndex = 0; CourseIndex < Students[0].Semesters[SemesterIndex].Courses.Count; CourseIndex++)
                {
                    Console.WriteLine("Semester {0} Course {1}, Course Average: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseAverages[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Variance: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseVariances[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Standard Deviation: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseStandardDeviations[CourseIndex]);
                    Console.WriteLine();
                }
            }
        }
    }
    public void PLOSpecificAllDecider()
    {
        int UserResponse;
        Console.WriteLine("Do you wish to output the selected statistics for all semesters or for specific semesters");
        UserResponse = Int32.Parse(Console.ReadLine());
    }
    // Possibly a method that outputs both PLO and semester statistics for all and selected semesters.
    // Make sure to go to a separate menu for outputing statistics. Separate menu navigation for it.
    // Or could diverge at each option and ask if they want it for specific semesters or for all semesters.
    public void PrintAllStatisticsAllSemesters()
    {
        PrintAllPLOAllSemesterStatistics();
        PrintAllSemesterCourseStatistics();
    }
    public void PrintAllStatisticsSpecificSemesters()
    {
        List<int> SelectedSem = new List<int>();
        SelectedSem.AddRange(SelectSpecificSemester());
        for (int PLOIndex = 0; PLOIndex < 12; PLOIndex++)
        {
            for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
            {
                var statistic = PLOStatistics[PLOIndex].PLOSemesterStatistics[SelectedSem[semIndex] - 1];
                for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
                {
                    Console.WriteLine("PLO {0} Statistics for Semester {1}", PLOIndex + 1, SelectedSem[semIndex]);
                    Console.WriteLine();
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Average CLO Marks are {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.AverageCLOMarks);
                    Console.WriteLine("For PLO {0}, in Semester {1}, Variance in Average CLO Marks is {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.Variance);
                    Console.WriteLine("For PLO {0}, in Semester {1}, the Standard Deviation in Average CLO Marks is {2}", PLOIndex + 1, SelectedSem[semIndex], statistic.StandardDeviation);
                    Console.WriteLine();
                }
            }
        }
        for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
        {
            for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
            {
                int currentSem = SelectedSem[SemesterIndex] - 1;
                Console.WriteLine("Course Statistics for Semester {0}", SemesterIndex + 1);
                Console.WriteLine();
                for (int CourseIndex = 0; CourseIndex < Students[0].Semesters[SemesterIndex].Courses.Count; CourseIndex++)
                {
                    Console.WriteLine("Semester {0} Course {1}, Course Average: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseAverages[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Variance: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseVariances[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Standard Deviation: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseStandardDeviations[CourseIndex]);
                    Console.WriteLine();
                }
            }
        }
    }
    public void PrintSpecificPLOAllCourseSpecificSemesterStatistics()
    {
        List<int> SelectedPLOs = new List<int>();
        List<int> SelectedSem = new List<int>();
        SelectedSem.AddRange(SelectSpecificSemester());
        SelectedPLOs.AddRange(SelectSpecificPLOs());
        for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
        {
            for (int SemesterIndex = 0; SemesterIndex < 8; SemesterIndex++)
            {
                int currentSem = SelectedSem[SemesterIndex] - 1;
                Console.WriteLine("Course Statistics for Semester {0}", SemesterIndex + 1);
                Console.WriteLine();
                for (int CourseIndex = 0; CourseIndex < Students[0].Semesters[SemesterIndex].Courses.Count; CourseIndex++)
                {
                    Console.WriteLine("Semester {0} Course {1}, Course Average: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseAverages[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Variance: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseVariances[CourseIndex]);
                    Console.WriteLine("Semester {0} Course {1}, Course Standard Deviation: {2}", currentSem + 1, CourseIndex + 1, SemesterCourseStatistics[currentSem].CourseStandardDeviations[CourseIndex]);
                    Console.WriteLine();
                }
            }
        }
    }
    public void PLOStatisticCalculations()
    {
        for (int PLOStatIndex = 0; PLOStatIndex < 12; PLOStatIndex++)
        {
            PLOStatistic PLOStatObject = new PLOStatistic();
            PLOStatObject.PLOSemesterStatistics = new List<PLOSemesterStatistic>();
            for (int PLOSemesterStatisticIndex = 0; PLOSemesterStatisticIndex < 8; PLOSemesterStatisticIndex++)
            {
                var ploSemStat = new PLOSemesterStatistic();
                PLOStatObject.PLOSemesterStatistics.Add(ploSemStat);
                PLOStatObject.PLOSemesterStatistics[PLOSemesterStatisticIndex].CLOMarks = new List<int>();
            }
            for (int StudentIndex = 0; StudentIndex < Students.Count; StudentIndex++)
            {
                for (int SemesterIndex = 0; SemesterIndex < Students[StudentIndex].Semesters.Count; SemesterIndex++)
                {
                    for (int CourseIndex = 0; CourseIndex < Students[StudentIndex].Semesters[SemesterIndex].Courses.Count; CourseIndex++)
                    {
                        PLOStatObject.PLOSemesterStatistics[SemesterIndex].TotalCLOMarks += Students[StudentIndex].Semesters[SemesterIndex].Courses[CourseIndex].PLOsCourses[PLOStatIndex].PLOTotalCourseMarks;
                        PLOStatObject.PLOSemesterStatistics[SemesterIndex].TotalCLOs += Students[StudentIndex].Semesters[SemesterIndex].Courses[CourseIndex].PLOsCourses[PLOStatIndex].PLOTotalCourseLinks;
                        PLOStatObject.PLOSemesterStatistics[SemesterIndex].CLOMarks.AddRange(Students[StudentIndex].Semesters[SemesterIndex].Courses[CourseIndex].PLOsCourses[PLOStatIndex].CLOMarksforDeviations);
                    }
                }
            }
            for (int SemesterStatIndex = 0; SemesterStatIndex < PLOStatObject.PLOSemesterStatistics.Count; SemesterStatIndex++)
            {
                if (PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOs != 0)
                {
                    PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].AverageCLOMarks = PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOMarks / PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOs;
                }
                for (int SumOfDeviationIndex = 0; SumOfDeviationIndex < PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].CLOMarks.Count; SumOfDeviationIndex++)
                {
                    PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].SumofDeviationsSquared = PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].SumofDeviationsSquared + ((PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].CLOMarks[SumOfDeviationIndex] - PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].AverageCLOMarks) * (PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].CLOMarks[SumOfDeviationIndex] - PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].AverageCLOMarks));
                }
                if (PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOs != 0 && PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].SumofDeviationsSquared != 0)
                {
                    PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].StandardDeviation = Math.Round(Math.Sqrt(PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].SumofDeviationsSquared / (PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOs - 1)),2);
                    PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].Variance = Math.Round((PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].SumofDeviationsSquared / (PLOStatObject.PLOSemesterStatistics[SemesterStatIndex].TotalCLOs - 1)),2);
                }
            }
            PLOStatistics.Add(PLOStatObject);
        }
    }
    public int MenuNavigationConsole()
    {
        int UserOption;
        bool falseInputInt = true;
        Console.WriteLine();
        Console.WriteLine("Data has been successfully entered. Please select further options.");
        Console.WriteLine("Enter 1 to print results for a specific student.");
        Console.WriteLine("Enter 2 to print results for all students.");
        Console.WriteLine("Enter 4 to receive the PLO description of a particular PLO.");
        Console.WriteLine("Enter 5 to receive the calculated PLO Statistics for all PLOs across all semesters.");
        Console.WriteLine("Enter 6 to receive the PLO statistics for all PLOs across specific semesters.");
        Console.WriteLine("Enter 7 to receive the PLO statistics for specific PLOs across all semesters.");
        Console.WriteLine("Enter 8 to receive the PLO statistics for specific PLOs across specific semesters.");
        Console.WriteLine("Enter -1 to return to input method selection.");
        falseInputInt = true;
        while (!Int32.TryParse(Console.ReadLine(), out UserOption) || falseInputInt)
        {
            if ((UserOption == 1 || UserOption == 2) || (UserOption == -1 || UserOption == 4) || (UserOption == 5 || UserOption == 6) || (UserOption == 7 || UserOption == 8))
            {
                falseInputInt = false;
                break;
            }
            Console.WriteLine("Incorrect Input. Please select one of the defined options");
            Console.WriteLine("1 to print results for all students, 2 to print results for all students");
            Console.WriteLine("4 to receive the description of a particular PLO, -1 to return to input method selection.");
            Console.WriteLine("Enter 5 to receive the calculated PLO Statistics for all PLOs across all semesters.");
            Console.WriteLine("Enter 6 to receive the PLO statistics for all PLOs across specific semesters.");
            Console.WriteLine("Enter 7 to receive the PLO statistics for specific PLOs across all semesters.");
            Console.WriteLine("Enter 8 to receive the PLO statistics for specific PLOs across specific semesters.");
        }
        return UserOption;
    }
    public int MenuNavigation()
    {
        int UserOption;
        
        Console.WriteLine();
        Console.WriteLine("Data has been successfully entered. Please select further options.");
        Console.WriteLine("Enter 1 to print results for a specific student");
        Console.WriteLine("Enter 2 to print results for all students");
        Console.WriteLine("Enter 3 to output results to an excel file");
        Console.WriteLine("Enter 4 to receive the description for a particular PLO.");
        Console.WriteLine("Enter 5 to receive the calculated PLO Statistics for all PLOs across all semesters.");
        Console.WriteLine("Enter 6 to receive the PLO statistics for all PLOs across specific semesters.");
        Console.WriteLine("Enter 7 to receive the PLO statistics for specific PLOs across all semesters.");
        Console.WriteLine("Enter 8 to receive the PLO statistics for specific PLOs across specific semesters.");
        Console.WriteLine("Enter -1 to return to input method selection");
        UserOption = OutputExcel.outputagain;
        
             
      
        return UserOption;
    }
    public void PrintAllStudents()
    {
        for (int stIndex = 0; stIndex < Students.Count; stIndex++)
        {
            var student = Students[stIndex];
            Console.WriteLine("Student name: " + student.Name);
            Console.WriteLine("Student Registration Number: " + student.StudentID);
            Console.WriteLine();
            for (int semIndex = 0; semIndex < student.Semesters.Count; semIndex++)
            {
                var semester = student.Semesters[semIndex];
                Console.WriteLine("Semester number: " + semester.SemNum);
                Console.WriteLine();
                for (int SemPLOIndex = 0; SemPLOIndex < 12; SemPLOIndex++)
                {
                    var SemPLOStat = semester.PLOsSemester[SemPLOIndex];
                    if (SemPLOStat.PLOTotalSemesterLinks != 0)
                    {
                        Console.WriteLine("Semester {0} PLO {1}: {2}", semester.SemNum, SemPLOIndex + 1, SemPLOStat.PLOSemState);
                    }
                }

                for (int cIndex = 0; cIndex < semester.TotalCourses; cIndex++)
                {
                    var course = semester.Courses[cIndex];
                    Console.WriteLine("Course ID: " + course.CourseID);
                    Console.WriteLine("Course Name: " + course.CourseName);
                    for (int cloIndex = 0; cloIndex < course.CLOs.Count; cloIndex++)
                    {
                        var clo = course.CLOs[cloIndex];
                        //Console.WriteLine("CLO {0} Description: {1}", course.CLOs[cloIndex].CLOnumber, clo.CLOdescription);
                        Console.WriteLine("CLO {0} Marks: {1}", course.CLOs[cloIndex].CLOnumber, clo.CLOmarks);
                        if (course.CLOs[cloIndex].PLOPass == 1)
                        {
                            Console.WriteLine("CLO {0} Status: Pass", course.CLOs[cloIndex].CLOnumber);
                        }
                        else
                        {
                            Console.WriteLine("CLO {0} Pass Status: Fail", course.CLOs[cloIndex].CLOnumber);

                        }
                    }
                    for (int CoursePLOIndex = 0; CoursePLOIndex < 12; CoursePLOIndex++)
                    {
                        var CoursePLOStat = course.PLOsCourses[CoursePLOIndex];
                        if (CoursePLOStat.PLOTotalCourseLinks != 0)
                        {
                            Console.WriteLine("PLO Course State for PLO {0}: {1}", CoursePLOIndex + 1, CoursePLOStat.PLOCourseState);
                        }
                    }
                    Console.WriteLine();
                }

            }

        }
    }
    public void PrintStudent()
    {
        //bool falseIntInput;
        //int InputChecker = 0;
        int studentMatch;
        //string UserResponse;
        List<int> SelectedSem = new List<int>();
        Console.WriteLine("Please enter the Registration Number of the student");
        while (!Int32.TryParse(Console.ReadLine(), out studentMatch))
        {
            Console.WriteLine("Incorrect Input. The Registration Number can only comprise digits from 0-9.");
            Console.Write("Enter the student's Registration Number again: ");
        }
        for (int stIndex = 0; stIndex < Students.Count; stIndex++)
        {
            if (Students[stIndex].StudentID == studentMatch)
            {
                studentMatch = stIndex;
                break;
            }
        }
        var student = Students[studentMatch];
        SelectedSem.AddRange(SelectSpecificSemester()); // DO NOT KNOW IF ASSIGNMENT POSSIBLE.
        Console.WriteLine("Student name: " + Students[studentMatch].Name);
        //bool PrintMoreSemesters = true;
        //Console.Write("For which semester(s) do you wish to output data. Please enter the semester's number: ");
        //while (PrintMoreSemesters)
        //{
        //    falseIntInput = true;
        //    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
        //    {
        //        if (1 <= InputChecker || InputChecker <= 8)
        //        {
        //            falseIntInput = false;
        //            break;
        //        }
        //        Console.WriteLine("Incorrect Input.Please enter a number from 1 to 8");
        //    }
        //    SelectedSem.Add(InputChecker);
        //    Console.Write("Do you wish to add another semester: (y/n) ");
        //    UserResponse = Console.ReadLine();
        //    while (UserResponse != "y" && UserResponse != "n")
        //    {
        //        Console.WriteLine("Incorrect Input. Please enter y to select another semester or n to proceed further: ");
        //        UserResponse = Console.ReadLine();
        //    }
        //    if (UserResponse == "n")
        //    {
        //        PrintMoreSemesters = false;
        //    }
        //}
        //SelectedSem = SelectedSem.OrderBy(x => x).ToList();
        for (int semIndex = 0; semIndex < SelectedSem.Count; semIndex++)
        {
            var semester = student.Semesters[SelectedSem[semIndex] - 1];
            Console.WriteLine("Semester Number: " + semester.SemNum);
            for (int SemPLOIndex = 0; SemPLOIndex < 12; SemPLOIndex++)
            {
                var SemPLOStat = semester.PLOsSemester[SemPLOIndex];
                if (SemPLOStat.PLOTotalSemesterLinks != 0)
                {
                    Console.WriteLine("Semester {0} PLO {1}: {2}", semester.SemNum, SemPLOIndex + 1, SemPLOStat.PLOSemState);
                }
            }
            for (int cIndex = 0; cIndex < semester.TotalCourses; cIndex++)
            {
                var course = semester.Courses[cIndex];
                Console.WriteLine("Course ID: " + course.CourseID);
                Console.WriteLine("Course Name: " + course.CourseName);
                for (int cloIndex = 0; cloIndex < course.CLOs.Count; cloIndex++)
                {
                    var clo = course.CLOs[cloIndex];
                    //Console.WriteLine("CLO {0} Description: {1}", course.CLOs[cloIndex].CLOnumber, clo.CLOdescription);
                    Console.WriteLine("CLO {0} Marks: {1} ", course.CLOs[cloIndex].CLOnumber, clo.CLOmarks);
                    if (course.CLOs[cloIndex].PLOPass == 1)
                    {
                        Console.WriteLine("CLO {0} Pass State: Pass", course.CLOs[cloIndex].CLOnumber);
                    }
                    else
                    {
                        Console.WriteLine("CLO {0} Pass State: Fail", course.CLOs[cloIndex].CLOnumber);

                    }
                }
                for (int CoursePLOIndex = 0; CoursePLOIndex < 12; CoursePLOIndex++)
                {
                    var CoursePLOStat = course.PLOsCourses[CoursePLOIndex];
                    if (CoursePLOStat.PLOTotalCourseLinks != 0)
                    {
                        Console.WriteLine("PLO State for Course {0}: {1}", course.CourseID, CoursePLOStat.PLOCourseState); // Should use placeholder and then add 1.
                    }
                }
            }

        }
    }
    public static List<int> SelectSpecificSemester()
    {
        int InputChecker = 0;
        bool falseIntInput;
        bool PrintMoreSemesters = true;
        string UserResponse;
        List<int> SelectedSem = new List<int>();
        Console.Write("For which semester(s) do you wish to output data. Please enter the semester's number: ");
        while (PrintMoreSemesters)
        {
            falseIntInput = true;
            while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
            {
                if (1 <= InputChecker || InputChecker <= 8)
                {
                    falseIntInput = false;
                    break;
                }
                Console.WriteLine("Incorrect Input.Please enter a number from 1 to 8");
            }
            SelectedSem.Add(InputChecker);
            Console.Write("Do you wish to add another semester: (y/n) ");
            UserResponse = Console.ReadLine();
            while (UserResponse != "y" && UserResponse != "n")
            {
                Console.WriteLine("Incorrect Input. Please enter y to select another semester or n to proceed further: ");
                UserResponse = Console.ReadLine();
            }
            if (UserResponse == "n")
            {
                PrintMoreSemesters = false;
            }
        }
        SelectedSem = SelectedSem.OrderBy(x => x).ToList();
        return SelectedSem;
    } //MUST TEST AND THEN REMOVE EXCESS CODE
    public static List<int> SelectSpecificPLOs() //MUST TEST AND THEN REMOVE EXCESS CODE
    {
        int InputChecker = 0;
        bool PrintMorePLOs = true;
        bool falseIntInput;
        string UserResponse;
        List<int> SelectedPLOs = new List<int>();
        Console.WriteLine("For which PLOs do you wish to output data. Please enter the PLO's number: ");
        while (PrintMorePLOs)
        {
            falseIntInput = true;
            while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
            {
                if (1 <= InputChecker || InputChecker <= 12)
                {
                    falseIntInput = false;
                    break;
                }
                Console.WriteLine("Incorrect Input.Please enter a number from 1 to 12");
            }
            SelectedPLOs.Add(InputChecker);
            Console.Write("Do you wish to add another PLO.");
            UserResponse = Console.ReadLine();
            while (UserResponse != "y" && UserResponse != "n")
            {
                Console.WriteLine("Incorrect Input. Please enter y to add another PLO or n to proceed further: ");
                UserResponse = Console.ReadLine();
            }
            if (UserResponse == "n")
            {
                PrintMorePLOs = false;
            }
        }
        SelectedPLOs = SelectedPLOs.OrderBy(x => x).ToList();
        return SelectedPLOs;
    }
    public void ConstructStudentExcel()
    {
        int StudentNameCounter = Mainform1.StudentNamesCount;
        for (int XMLIndex = 0; XMLIndex < StudentNameCounter; XMLIndex++)
        {
            Student student = new Student();
            student.StudentID = (int)Mainform1.Studentids[XMLIndex];
            student.Name = Mainform1.studentnames[XMLIndex];
            int numberofCLOs = 0;
            for (int XMLSemesterIndex = 0; XMLSemesterIndex < Mainform1.NumberOfSemesters; XMLSemesterIndex++) //update form 1 code.
            {
                Semester sem = new Semester();
                sem.SemNum = XMLSemesterIndex + 1;
                sem.PLOsSemester = new List<PLO>();
                for (int PLOsSemesterIndex = 0; PLOsSemesterIndex < 12; PLOsSemesterIndex++)
                {
                    var plo = new PLO();

                    sem.PLOsSemester.Add(plo);
                }
                sem.TotalCourses = Mainform1.semesterlist1[XMLSemesterIndex].Count - 1;
                for (int CourseIndex = 1; CourseIndex < Mainform1.semesterlist1[XMLSemesterIndex].Count; CourseIndex++)
                {
                    Course course = new Course();
                    course.CourseID = Mainform1.semesterlist1[XMLSemesterIndex][CourseIndex];
                    course.PLOsCourses = new List<PLO>();
                    for (int PLOsCoursesIndex = 0; PLOsCoursesIndex < 12; PLOsCoursesIndex++)
                    {
                        var plo = new PLO();
                        if (plo.CLOMarksforDeviations == null)
                        {
                            plo.CLOMarksforDeviations = new List<int>();
                        }
                        course.PLOsCourses.Add(plo);
                    }

                    for (int CLOIndex = 0; CLOIndex < Int32.Parse(Mainform1.CLOs[numberofCLOs][1]); CLOIndex++)
                    {
                        CLO clo = new CLO();
                        clo.CLOnumber = CLOIndex;
                        clo.CLOmarks = (int)Mainform1.studentdata[XMLSemesterIndex][CourseIndex - 1][XMLIndex][CLOIndex][0];
                        if (clo.CLOmarks >= 50)
                        {
                            clo.PLOPass = 1;
                        }
                        else
                        {
                            clo.PLOPass = 0;
                        }
                        clo.PLOLink = (int)Mainform1.studentdata[XMLSemesterIndex][CourseIndex - 1][XMLIndex][CLOIndex][1];
                        if (clo.PLOLink != 0)
                        {
                            course.TotalCourseMarks += clo.CLOmarks;
                            course.TotalCLOs += 1;
                            course.PLOsCourses[clo.PLOLink - 1].CLOMarksforDeviations.Add(clo.CLOmarks);
                            course.PLOsCourses[clo.PLOLink - 1].PLOTotalCourseMarks += clo.CLOmarks;
                            course.PLOsCourses[clo.PLOLink - 1].PLOTotalCourseLinks += 1;
                            if (course.CLOs == null)
                            {
                                course.CLOs = new List<CLO>();
                            }
                            course.CLOs.Add(clo);
                        }

                    }
                    numberofCLOs++;
                    for (int CoursePLOIndex = 0; CoursePLOIndex < 12; CoursePLOIndex++)
                    {
                        if (course.PLOsCourses[CoursePLOIndex].PLOTotalCourseLinks != 0)
                        {
                            sem.PLOsSemester[CoursePLOIndex].PLOTotalSemesterLinks += 1;
                            if (course.PLOsCourses[CoursePLOIndex].PLOTotalCourseMarks >= (50 * course.PLOsCourses[CoursePLOIndex].PLOTotalCourseLinks))
                            {
                                course.PLOsCourses[CoursePLOIndex].PLOCourseState = "Pass";
                                sem.PLOsSemester[CoursePLOIndex].PLOTotalSemesterPasses += 1;
                            }
                            else { course.PLOsCourses[CoursePLOIndex].PLOCourseState = "Fail"; }
                        }
                    }
                    if (sem.Courses == null)
                    {
                        sem.Courses = new List<Course>();
                    }
                    course.CourseAverage = course.TotalCourseMarks / course.TotalCLOs;
                    sem.Courses.Add(course);
                }

                for (int PLOsSemesterIndex = 0; PLOsSemesterIndex < 12; PLOsSemesterIndex++)
                {
                    if (sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks == 0)
                    {
                        sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Unevaluated";
                    }
                    else if ((sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterPasses >= (0.8 * sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks)) && (sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks != 0))
                    {
                        sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Pass";
                    }
                    else
                    {
                        sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Fail";
                    }
                }
                if (student.Semesters == null)
                {
                    student.Semesters = new List<Semester>();
                }
                student.Semesters.Add(sem);
            }
            Students.Add(student);
        }
    }


    public Student ConstructStudent()
    {
        string UserResponse;
        int InputChecker;
        bool falseIntInput;
        Student student = new Student();
        Console.Write("Please enter the student's Registration Number: ");
        while (!Int32.TryParse(Console.ReadLine(), out InputChecker))
        {
            Console.WriteLine("Incorrect Input. The Registration Number can only comprise digits from 0-9.");
            Console.Write("Enter the student's Registration Number again: ");
        }
        student.StudentID = InputChecker;
        Console.Write("Enter the name of the student: ");
        student.Name = Console.ReadLine();
        Console.WriteLine();

        bool addMoreSemesters = true;
        while (addMoreSemesters)
        {
            Semester sem = new Semester();
            Console.Write("Enter the semester number: ");
            falseIntInput = true;
            while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
            {
                if (1 <= InputChecker && InputChecker <= 8)
                {
                    falseIntInput = false;
                    break;
                }
                Console.WriteLine("Semester number must be an integer from 1 to 8");
                Console.WriteLine("Please enter the semester number again");
            }
            sem.SemNum = InputChecker;
            // Ask for Semester INformation

            // Add Courses
            sem.PLOsSemester = new List<PLO>();
            for (int PLOsSemesterIndex = 0; PLOsSemesterIndex < 12; PLOsSemesterIndex++)
            {
                var plo = new PLO();
                sem.PLOsSemester.Add(plo);
            }
            bool addMoreCourses = true;
            while (addMoreCourses)
            {
                sem.TotalCourses += 1;
                Course course = new Course();
                Console.Write("Please enter the Course ID: ");
                course.CourseID = Console.ReadLine();
                Console.Write("Please enter the Course Name: ");
                course.CourseName = Console.ReadLine();
                course.PLOsCourses = new List<PLO>();
                for (int PLOsCoursesIndex = 0; PLOsCoursesIndex < 12; PLOsCoursesIndex++)
                {
                    var plo = new PLO();
                    if (plo.CLOMarksforDeviations == null)
                    {
                        plo.CLOMarksforDeviations = new List<int>();
                    }
                    course.PLOsCourses.Add(plo);
                }
                bool addMoreCLOs = true;
                int cloCounter = 1;
                while (addMoreCLOs)
                {
                    CLO clo = new CLO();
                    clo.CLOnumber = cloCounter;
                    Console.Write("Please enter the marks for CLO {0} (from 0-100): ", cloCounter);
                    falseIntInput = true;
                    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
                    {
                        if (0 <= InputChecker && InputChecker <= 100)
                        {
                            falseIntInput = false;
                            break;
                        }
                        Console.WriteLine("Incorrect Input. Please enter an integer from 0 to 100");
                    }
                    if (clo.CLOmarks >= 50)
                    {
                        clo.PLOPass = 1;
                    }
                    else
                    {
                        clo.PLOPass = 0;
                    }
                    Console.Write("Please enter the number of the PLO to which CLO {0} is linked (1-12): ", cloCounter);
                    falseIntInput = true;
                    while (!Int32.TryParse(Console.ReadLine(), out InputChecker) || falseIntInput)
                    {
                        if (1 <= InputChecker && InputChecker <= 12)
                        {
                            falseIntInput = false;
                            break;
                        }
                        Console.WriteLine("Incorrect Input. Please enter a number from 1 (inclusive) to 12 (inclusive) to select a PLO");
                    }
                    course.TotalCourseMarks += clo.CLOmarks;
                    course.TotalCLOs += 1;
                    clo.PLOLink = InputChecker - 1;
                    course.PLOsCourses[clo.PLOLink].CLOMarksforDeviations.Add(clo.CLOmarks);
                    course.PLOsCourses[clo.PLOLink].PLOTotalCourseLinks += 1;
                    course.PLOsCourses[clo.PLOLink].PLOTotalCourseMarks += clo.CLOmarks;
                    if (course.CLOs == null)
                    {
                        course.CLOs = new List<CLO>();
                    }
                    course.CLOs.Add(clo);
                    Console.Write("Do you wish to enter another CLO (y/n): ");
                    UserResponse = Console.ReadLine();
                    while (UserResponse != "y" && UserResponse != "n")
                    {
                        Console.Write("Incorrect Input. Please select either y to add another CLO or n to not add another CLO: ");
                        UserResponse = Console.ReadLine();
                    }
                    if (UserResponse == "n")
                    {
                        addMoreCLOs = false;
                    }
                    cloCounter++;
                }
                for (int CoursePLOIndex = 0; CoursePLOIndex < 12; CoursePLOIndex++)
                {
                    if (course.PLOsCourses[CoursePLOIndex].PLOTotalCourseLinks != 0)
                    {
                        sem.PLOsSemester[CoursePLOIndex].PLOTotalSemesterLinks += 1;
                        if (course.PLOsCourses[CoursePLOIndex].PLOTotalCourseMarks >= (50 * course.PLOsCourses[CoursePLOIndex].PLOTotalCourseLinks))
                        {
                            course.PLOsCourses[CoursePLOIndex].PLOCourseState = "Pass";
                            sem.PLOsSemester[CoursePLOIndex].PLOTotalSemesterPasses += 1;
                        }
                        else { course.PLOsCourses[CoursePLOIndex].PLOCourseState = "Fail"; }
                    }
                }
                course.CourseAverage = course.TotalCourseMarks / course.TotalCLOs;
                if (sem.Courses == null)
                {
                    sem.Courses = new List<Course>();
                }

                sem.Courses.Add(course);

                Console.Write("Do you want to add another course? (y/n) ");
                UserResponse = Console.ReadLine();
                if (UserResponse == "n")
                {
                    addMoreCourses = false;
                }
            }
            for (int PLOsSemesterIndex = 0; PLOsSemesterIndex < 12; PLOsSemesterIndex++)
            {
                if (sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks == 0)
                {
                    sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Unevaluated";
                }
                else if ((sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterPasses >= (0.8 * sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks)) && (sem.PLOsSemester[PLOsSemesterIndex].PLOTotalSemesterLinks != 0))
                {
                    sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Pass";
                }
                else
                {
                    sem.PLOsSemester[PLOsSemesterIndex].PLOSemState = "Fail";
                }
            }
            if (student.Semesters == null)
            {
                student.Semesters = new List<Semester>();
            }
            student.Semesters.Add(sem);
            Console.WriteLine();
            Console.Write("Do you want to add another semester? (y/n)");
            UserResponse = Console.ReadLine();
            Console.WriteLine();
            if (UserResponse == "n")
            {
                addMoreSemesters = false;
            }
        }
        return student;
    }

    public class Program
    {
        //the Main method, which is the insertion point for the program.
        public static void Main(string[] args)
        {
            var manager = new StudentManager();
            manager.Start();
        }
    }

    public void printPLO(int plonumber)
    {
        //switch (plonumber)
        //{
           
        //}

    }
}