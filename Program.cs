using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CPDData1
{
    class Program
    {
        public static List<DateTime> StartdateList = new List<DateTime>();
        public static List<DateTime> EnddateList = new List<DateTime>();
        public static List<double> averagesList=new List<double>();
        static void Main(string[] args)
        {
            DataTable dt = readFile();
            isValidData(dt);
            calculateAverages(dt);
        }

        public static DataTable readFile()
        {
            DataTable firstTable = new DataTable();
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String[] excelSheets;
            string excelFile = "Target CPD.xlsx";
            try
            {
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=Excel 12.0 xml;";
                objConn = new OleDbConnection(connString);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    // return null;
                    Console.WriteLine("No data found");
                }
                excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                String query = "SELECT MissionStart,MIssionStop,ProxyCPDData FROM [" + excelSheets[0] + "]"; // You can use any different queries to get the data from the excel sheet
                OleDbConnection conn = new OleDbConnection(connString);

                cmd = new OleDbCommand(query, conn);
                da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                firstTable = ds.Tables[0];
                // return firstTable;
                return firstTable;
            }
            catch (Exception exc)
            {

            }

            return firstTable;
        }

        public static bool isValidData(DataTable dt)
        {
           // bool isDataValid = true;
            //int count = 0;
            for(int i=0;i<dt.Rows.Count-1;i++)
            {
                DataRow dr1 = dt.Rows[i];
                DataRow dr2 = dt.Rows[i + 1];
                DateTime endDay = DateTime.Parse(dr1[1].ToString());
                DateTime nextBeginDay = DateTime.Parse(dr2[0].ToString());
                double dayDifference = endDay.Subtract(nextBeginDay).TotalDays;
                if (dayDifference < 0.0)
                {
                    Console.WriteLine("Error in line: " + i);
                    Console.ReadKey();
                    return false ;
                }                  
            }
            //Console.WriteLine("Read Successful");
           // Console.ReadKey();
            return true;
        }

        //public static void calculateAverages(DataTable dt)
        //{
        //    DateTime StartDay, EndDay;
        //    foreach(DataRow dr in dt.Rows)
        //    {
        //        //GEt the first day and last day
        //        StartDay = Convert.ToDateTime(dr[0].ToString());
        //        EndDay = Convert.ToDateTime(dr[1].ToString());

        //        //Get the total number of days
        //        double TotalDays = Math.Ceiling((EndDay.Date - StartDay.Date).TotalDays+1);

        //        //Clculate the average CPD for the given period
        //        double AvgCPD = Convert.ToDouble(dr[2])/ TotalDays;

        //        //Find the number of days in the month of the Startday and calculate averge
        //        double firstmonthAvg = AvgCPD * DateTime.DaysInMonth(StartDay.Year, StartDay.Month);
        //        averagesList.Add(firstmonthAvg);
        //        dateList.Add(new DateTime(StartDay.Year, StartDay.Month,1));
                

        //        //Calculate the months for the months between startDay and end day
        //        double remaningMonths = EndDay.Month - StartDay.Month;
        //        remaningMonths = remaningMonths < 0 ? 12 + remaningMonths : remaningMonths;
        //        if(remaningMonths>1)
        //        {
        //            for(int i=1;i<remaningMonths;i++)
        //            {
        //                DateTime tempDay = StartDay.AddMonths(i);
        //                dateList.Add(new DateTime(tempDay.Year,tempDay.Month,1));
        //                averagesList.Add(AvgCPD*DateTime.DaysInMonth(tempDay.Year,tempDay.Month));
        //            }
        //        }

        //        //Find the number of days in the month of the Endday and calculate averge
        //        double lastMonthAvg = AvgCPD * DateTime.DaysInMonth(EndDay.Year, EndDay.Month); //The number of days should be inclusive
        //        dateList.Add(new DateTime(EndDay.Year, EndDay.Month, 1));
        //        averagesList.Add(lastMonthAvg);
        //    }
        //}
    
    public static void calculateAverages(DataTable dt)
    {
        DateTime StartDay, EndDay;
            foreach(DataRow dr in dt.Rows)
            {
                //GEt the first day and last day
                StartDay = Convert.ToDateTime(dr[0].ToString());
                EndDay = Convert.ToDateTime(dr[1].ToString());

                //Get the total number of days
                double TotalDays = Math.Ceiling((EndDay.Date - StartDay.Date).TotalDays+1);

                //Calculate the average CPD for the given period
                double AvgCPD = Convert.ToDouble(dr[2])/ TotalDays;

                //Add the average CPD, first month and last month values to the list
                averagesList.Add(AvgCPD);
                StartdateList.Add(new DateTime(StartDay.Year, StartDay.Month, 1));
                EnddateList.Add(new DateTime(EndDay.Year, EndDay.Month, 1));


            }
    }
    
    }
}
