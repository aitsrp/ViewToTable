using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Web;
using System.IO;

namespace ViewToTable
{
    public class MySQLServer
    {
        private static string ConnectionString;

        public static string Driver = "{MySQL ODBC 5.3 Unicode Driver}";

        public static string Name = "192.168.43.85";

        public static string UserName = "root";

        public static string Password = "Wgraphics";

        public static string Database = "votf_reports";

        public string Error { get; set; }

        public string Query { get; set; }

        public int Row { get; set; }

        public string Filename { get; set; }

        public void Initialize()
        {
            ConnectionString = String.Format("Driver={0}; Server={1}; User={2}; password={3}; Database={4}; Option=3;", Driver, Name, UserName, Password, Database);
        }

        public DataTable ExecuteQuery()
        {
            //Reference: http://msdn.microsoft.com/en-us/library/ms998569.aspx
            //When Using DataReaders, Specify CommandBehavior.CloseConnection
            DataTable dReturnValue = new DataTable();

            try
            {
                //Open and Close the Connection in the Method. See Reference.
                OdbcDataAdapter MySQLAdapter = new OdbcDataAdapter(Query, ConnectionString);

                //Explicitly Close Connections. See Reference.
                MySQLAdapter.SelectCommand.Connection.Close();

                //Do Not Explicitly Open a Connection if You Use Fill or Update for a Single Operation
                MySQLAdapter.Fill(dReturnValue);

                Error = "";
            }
            catch (Exception e)
            {
                Error = e.Message;
                WriteToFile();
                Console.WriteLine("Failed to execute: '" + Query + "' with error: " + Error);
            }

            return dReturnValue;
        }

        public bool ExecuteNonQuery()
        {
            OdbcConnection MySQLConnection = new OdbcConnection(ConnectionString);
            try
            {
                //Open and Close the Connection in the Method. See Reference.
                MySQLConnection.Open();
                OdbcCommand MySQLCommand = new OdbcCommand(Query, MySQLConnection);
                MySQLCommand.ExecuteNonQuery();

                Error = "";
                return true;
            }
            catch (Exception e)
            {
                Error = e.Message;
                WriteToFile();
                Console.WriteLine("ExecuteNonQuery: " + Query + "Error: " + Error);
                return false;
            }
            finally
            {
                //Open and Close the Connection in the Method. See Reference.
                MySQLConnection.Close();
            }
        }

        public void WriteToFile()
        {
            using (StreamWriter sw = new StreamWriter("D:\\Projects\\ViewToTable\\ViewToTableErrors.txt", true))
            {
                sw.WriteLine("Filename: " + Filename);
                sw.WriteLine("Row: " + Row);
                sw.WriteLine("Error: " + Error);
                sw.WriteLine("Query: " + Query);
                sw.WriteLine(" ");
            }
        }
    }
}
