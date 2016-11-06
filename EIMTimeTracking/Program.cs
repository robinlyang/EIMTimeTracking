using System;
using System.Collections;
using System.IO;

namespace EIMTimeTracking
{
    class Program
    {
        public static object MessageBox { get; private set; }

        static void Main(string[] args)
        {
            menu();
        }

        static void menu()
        {
            String startTimeStr = "";
            DateTime startDate = DateTime.Today;
            DateTime endTime = DateTime.Now;
            DateTime actualEndTime = DateTime.Now;
            DateTime startTime;
            int addHours = 0;
            int addMinutes = 0;
            TimeSpan addTime;
            int breakTime;
            TimeSpan minusTime;

            double totalHours = 0.0;

            String taskStr;
            ArrayList taskList = new ArrayList();

            Console.WriteLine("Start time (E.g., 0830): ");
            startTimeStr = Console.ReadLine();
            addHours = Convert.ToInt32(startTimeStr.Substring(0, 2));
            addMinutes = Convert.ToInt32(startTimeStr.Substring(2, 2));
            addTime = new TimeSpan(addHours, addMinutes, 0);
            startTime = startDate.Add(addTime);
            Console.WriteLine("startTime: " + startTime);
            Console.WriteLine("End Time will be: " + endTime);

            Console.WriteLine("Break time in minutes (E.g., 30): ");
            breakTime = Convert.ToInt32(Console.ReadLine());
            minusTime = new TimeSpan(0, breakTime, 0);
            endTime = endTime.Subtract(minusTime);
            TimeSpan hoursWorked = endTime - startTime;
            double hoursWorkedDecimal = hoursWorked.TotalHours;
            hoursWorkedDecimal = Math.Round(hoursWorkedDecimal, 2);
            Console.WriteLine("Total hours: " + hoursWorkedDecimal.ToString());

            Console.WriteLine("Enter tasks delimited by a /: ");
            taskStr = Console.ReadLine();
            char delimiter = '/';
            String[] substrings = taskStr.Split(delimiter);
            //
            //14-Sep-16, 0830, 1700, 30, , , ,Task 1/Task 2
            //Yang, September, 20/Sep/16, 0.00, , Task 1  << Line 1
            string lines = "EDO\r\n" + startDate.ToString("dd-MMM-yy")
                + "," + startTime.ToString("hhmm") 
                + "," + actualEndTime.ToString("hhmm") 
                + "," + breakTime 
                + ",,,," 
                + taskStr + "\r\n"
                + "TIME TRACKING" 
                + "\r\n"; 
            foreach (var substring in substrings)
            {
                Console.WriteLine(substring + " added");
                taskList.Add(substring);
            }
            for (int i = 0; i < taskList.Count; i++)
            {
                taskList[i] = "Yang,"
                    + startTime.ToString("MMMM")
                    + "," + startTime.ToString("dd/MMM/yy")
                    + ",,,,,"
                    + taskList[i] + "\r\n";
                lines = lines + taskList[i].ToString();
            }

            System.IO.StreamWriter file = new StreamWriter(@"c:\\Users\RobinY\Documents\time.csv");
            file.WriteLine(lines);
            file.Close();

            String linesTwo = ",,,," + hoursWorkedDecimal.ToString();
            System.IO.StreamWriter fileTwo = new StreamWriter(@"c:\\Users\RobinY\Documents\time.csv", true);
            fileTwo.WriteLine(linesTwo);
            fileTwo.Close();

            Console.ReadLine();
            //System.Diagnostics.Process.Start(@"O:\staff\Enterprise Information Management\data\Common\Administration\Time Away\0 EDO Tracking Sheets\Robin - EDO Tracking Sheet 2016.xlsx");
            //System.Diagnostics.Process.Start(@"O:\staff\Enterprise Information Management\data\Common\Administration\Time Tracking\Yang_Time Tracking.xlsx");

            //

            System.Diagnostics.Process.Start(@"c:\\Users\RobinY\Documents\time.txt");
            System.Diagnostics.Process.Start(@"c:\\Users\RobinY\Documents\time.csv");
            System.Diagnostics.Process.Start(@"c:\\Users\RobinY\Documents\TimeMACRO.xlsm","true");
            //

            //
            Console.ReadLine();


        }
    }
}
