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
        static void Main(string[] args)
        {
            DataTable dt = readFile();
            isValidData(dt);

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
            Console.WriteLine("Read Successfull");
            Console.ReadKey();
            return true;
        }
    }
}
