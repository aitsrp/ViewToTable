using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ViewToTable
{
    public partial class Form1 : Form
    {
        MySQLServer server = null;

        public Form1()
        {
            InitializeComponent();
            Initialize();
        }

        public void Initialize()
        {
            server = new MySQLServer();
            server.Initialize();
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            DragDropEffects effect = DragDropEffects.None;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                FileInfo inf = new FileInfo(path);
                string ext = inf.Extension.ToLower();
                if (ext == ".xls" || ext == ".xlsx" || ext == ".vls")
                    effect = DragDropEffects.Copy;
            }

            e.Effect = effect;
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var filelist = ((string[])e.Data.GetData(DataFormats.FileDrop));
                textBox1.Text = "";
                for (var i = 0; i < filelist.Length; i++)
                {
                    string path = filelist[i];
                    FileInfo inf = new FileInfo(path);
                    var filename = inf.Name;
                    progressBar1.Value = 0;
                    if(inf.Extension == ".vls")
                    {
                        List<string> viewname = new List<string>();
                        var completed = 0;

                        string[] lines = System.IO.File.ReadAllLines(path);
                        foreach (string line in lines)
                        {
                            if(line.StartsWith("//") || (line == ""))
                            {
                                Console.WriteLine("ignored: " + line);
                            }
                            else
                            {
                                if(line.Contains(", "))
                                {
                                    string[] arr = line.Split(',');
                                    viewname.AddRange(arr);
                                }
                                else
                                {
                                    viewname.Add(line);
                                }
                            }
                        }
                        foreach (string v in viewname)
                        {
                            Console.WriteLine("Processing view " + v.Trim());
                            Proj_LogError("Processing view " + v.Trim());
                            ViewToTable(v.Trim());
                            completed++;
                            progressBar1.Minimum = 0;
                            progressBar1.Maximum = viewname.Count;
                            progressBar1.PerformStep();
                            label1.Text = completed + " of " + viewname.Count + " views";
                            label1.Refresh();
                        }
                    }
                    else if(filename.Contains("234584")) //General Survey Results
                    {
                        Console.WriteLine("Processing General Survey Results...");
                        Proj_LogError("Processing General Survey Results...");
                        generalSurveyResults(path);
                    }
                    else if (filename.Contains("972221")) // Product Survey Results
                    {
                        Console.WriteLine("Processing Product Survey Results...");
                        Proj_LogError("Processing Product Survey Results...");
                        productSurveyResults(path);
                    }
                    else if (filename.Contains("977717")) //Retailer Survey Results
                    {
                        Console.WriteLine("Processing Retailer Survey Results...");
                        Proj_LogError("Processing Retailer Survey Results...");
                        retailerSurveyResults(path);
                    }
                    //textBox1.Text += ExcelToJSON(path);
                    if (i == filelist.Length - 1)
                    {
                        Console.WriteLine("Process done!");
                        Proj_LogError("Process done!");
                    }
                }
                if (textBox1.Text != "")
                    Clipboard.SetText(textBox1.Text);
            }
        }

        private void ViewToTable(string filename)
        {
            server.Filename = filename;
            string queryView = "SELECT * FROM " + filename + ";";
            server.Query = queryView;
            textBox1.Text = "queryView: " + queryView;
            //Console.WriteLine("queryView: " + queryView);//delete
            bool exists = server.ExecuteNonQuery();
            if (exists)
            {
                DataTable dt = server.ExecuteQuery();
                var tablename = getTableName(filename);
                List<string> cols = new List<string>();
                string queryCreate;
                queryCreate = "DROP TABLE IF EXISTS " + tablename + "; ";
                server.Query = queryCreate;
                server.ExecuteNonQuery();

                queryCreate = "CREATE TABLE `" + tablename + "` (";
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    if (c == (dt.Columns.Count - 1)) //if it's the last column
                    {
                        cols.Add(dt.Columns[c].ColumnName);
                        queryCreate += "`" + dt.Columns[c].ColumnName + "` text) ENGINE=InnoDB DEFAULT CHARSET=utf8;";
                    }
                    else //while it's not the last column
                    {
                        cols.Add(dt.Columns[c].ColumnName);
                        queryCreate += "`" + dt.Columns[c].ColumnName + "` text, ";
                    }
                }

                textBox1.Text = "queryCreate: " + queryCreate;
                //Console.WriteLine("queryCreate: " + queryCreate);//delete
                server.Query = queryCreate;
                server.ExecuteNonQuery();

                //get rows
                string queryInsert = "";
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    queryInsert = "INSERT INTO " + tablename + " (";
                    for (int s = 0; s < cols.Count; s++) //for column names
                    {
                        if (s == (cols.Count - 1)) //if it's the last column
                        {
                            queryInsert += cols[s] + ") VALUES (";
                        }
                        else //while it's not the last column
                        {
                            queryInsert += cols[s] + ",";
                        }
                    }
                    for (int t = 0; t < dt.Columns.Count; t++)
                    {
                        if (t == (dt.Columns.Count - 1)) //if it's the last column
                        {
                            queryInsert += "'" + dt.Rows[r][t].ToString().Replace("'", "''") + "');";
                        }
                        else //while it's not the last column
                        {
                            queryInsert += "'" + dt.Rows[r][t].ToString().Replace("'", "''") + "',";
                        }
                    }
                    textBox1.Text = "queryInsert: " + queryInsert;
                    //Console.WriteLine("queryInsert: " + queryInsert);//delete
                    server.Row = r;
                    server.Query = queryInsert;
                    server.ExecuteNonQuery();

                }
            }
            else
            {
                Console.WriteLine("View " + filename + " does not exist.");
                Proj_LogError("View " + filename + " does not exist.");
            }
        }

        private void generalSurveyResults(string file)
        {
            server.Filename = file;
            server.Query = "TRUNCATE _generalsurveyresults1;";
            server.ExecuteNonQuery();
            server.Query = "TRUNCATE _generalsurveyresults2;";
            server.ExecuteNonQuery();
            server.Query = "SELECT * FROM _generalsurveyresults1;";
            DataTable dt1 = server.ExecuteQuery();
            server.Query = "SELECT * FROM _generalsurveyresults2;";
            DataTable dt2 = server.ExecuteQuery();
            if (file != "")
            {
                //Read File
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);

                for (var i = 1; i <= xlWorkbook.Sheets.Count; i++)
                {
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[i];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    
                    for(var j=2; j < 334; j++)//loop through entries aka rows because excel rows start at row 2 and ends at row 333
                    {
                        //generalsurveyresults1
                        string queryInsert1 = "INSERT INTO _generalsurveyresults1 (";
                        for( var k=1; k<133; k++)//get column names for query
                        {
                            if (k == 132) //if it's the last column
                            {
                                queryInsert1 += "`" + xlRange.Cells[1, k].Value2.Trim() + "`) VALUES (";
                            }
                            else //while it's not the last column
                            {
                                queryInsert1 += "`" + xlRange.Cells[1, k].Value2.Trim() + "`,";
                            }
                        }
                        for (var m = 1; m < 133; m++)//get values for query
                        {
                            if (m == 132) //if it's the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert1 += "NULL);";
                                }
                                else
                                {
                                    bool result = dt1.Columns[(m - 1)].DataType == System.Type.GetType("System.String");
                                    if (result)
                                    {
                                        queryInsert1 += "'" + xlRange.Cells[j, m].Value2.ToString().Replace("'","\'") + "');";
                                    }
                                    else
                                        queryInsert1 += "'" + xlRange.Cells[j, m].Value2 + "');";
                                }
                            }
                            else //while it's not the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert1 += "NULL,";
                                }
                                else
                                {
                                    bool result = dt1.Columns[(m - 1)].DataType == System.Type.GetType("System.String");
                                    if (result)
                                    {
                                        queryInsert1 += "'" + xlRange.Cells[j, m].Value2.ToString().Replace("'", "\'") + "',";
                                    }
                                    else
                                        queryInsert1 += "'" + xlRange.Cells[j, m].Value2 + "',";

                                }
                            }
                        }
                        textBox1.Text = "queryInsert1: " + queryInsert1;
                        //Console.WriteLine("queryInsert1: " + queryInsert1);
                        server.Row = j;
                        server.Query = queryInsert1;
                        server.ExecuteNonQuery();

                        //generalsurveyresults2
                        string queryInsert2 = "INSERT INTO _generalsurveyresults2 (`" + xlRange.Cells[1, 1].Value2.Trim() + "`,";
                        for (var n = 133; n < 253; n++)//get column names for query
                        {
                            if (n == 252) //if it's the last column
                            {
                                queryInsert2 += "`" + xlRange.Cells[1, n].Value2.Trim() + "`) VALUES ('" + xlRange.Cells[j, 1].Value2 + "',";
                            }
                            else //while it's not the last column
                            {
                                queryInsert2 += "`" + xlRange.Cells[1, n].Value2.Trim() + "`,";
                            }
                        }
                        for (var o = 133; o < 253; o++)//get values for query
                        {
                            if (o == 252) //if it's the last column
                            {
                                if (xlRange.Cells[j, o].Value2 == null)
                                {
                                    queryInsert2 += "NULL);";
                                }
                                else
                                {

                                    bool result = dt2.Columns[(o - 133)].DataType == System.Type.GetType("System.String");
                                    if (result)
                                    {
                                        queryInsert2 += "'" + xlRange.Cells[j, o].Value2.ToString().Replace("'", "\'") + "');";
                                    }
                                    else
                                        queryInsert2 += "'" + xlRange.Cells[j, o].Value2 + "');";
                                }
                            }
                            else //while it's not the last column
                            {
                                if (xlRange.Cells[j, o].Value2 == null)
                                {
                                    queryInsert2 += "NULL,";
                                }
                                else
                                {
                                    //Console.WriteLine(dt2.Columns[(o - 1)].DataType == System.Type.GetType("System.String"));
                                    bool result = dt2.Columns[(o - 133)].DataType == System.Type.GetType("System.String");
                                    if (result)
                                    {
                                        queryInsert2 += "'" + xlRange.Cells[j, o].Value2.ToString().Replace("'", "\'") + "',";
                                    }
                                    else
                                        queryInsert2 += "'" + xlRange.Cells[j, o].Value2 + "',";
                                }
                            }
                        }
                        textBox1.Text = "queryInsert2: " + queryInsert2;
                        //Console.WriteLine("queryInsert2: " + queryInsert2);
                        server.Query = queryInsert2;
                        server.ExecuteNonQuery();

                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = 332;
                        progressBar1.PerformStep();
                        label1.Text = (j-1) + " of " + 332 + " rows";
                        label1.Refresh();
                    }


                    if (i == xlWorkbook.Sheets.Count)
                    {
                        Console.WriteLine("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                        Proj_LogError("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                    }
                }
            }

        }

        private void productSurveyResults(string file)
        {
            server.Filename = file;
            server.Query = "TRUNCATE _productsurveyresults";
            server.ExecuteNonQuery();
            if (file != "")
            {
                //Read File
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);

                for (var i = 1; i <= xlWorkbook.Sheets.Count; i++)
                {
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[i];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    for (var j = 2; j < 33; j++)//loop through entries aka rows
                    {
                        string queryInsert = "INSERT INTO _productsurveyresults (";
                        for (var k = 1; k < 19; k++)//get column names for query
                        {
                            if (k == 18) //if it's the last column
                            {
                                queryInsert += "`" + xlRange.Cells[1, k].Value2.Trim() + "`) VALUES (";
                            }
                            else //while it's not the last column
                            {
                                queryInsert += "`" + xlRange.Cells[1, k].Value2.Trim() + "`,";
                            }
                        }
                        for (var m = 1; m < 19; m++)//get values for query
                        {
                            if (m == 18) //if it's the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert += "NULL);";
                                }
                                else
                                {
                                    queryInsert += "'" + xlRange.Cells[j, m].Value2 + "');";
                                }
                            }
                            else //while it's not the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert += "NULL,";
                                }
                                else
                                {
                                    queryInsert += "'" + xlRange.Cells[j, m].Value2 + "',";
                                }
                                
                            }
                        }
                        textBox1.Text = "queryInsert: " + queryInsert;
                        //Console.WriteLine("queryInsert: " + queryInsert);
                        server.Row = j;
                        server.Query = queryInsert;
                        server.ExecuteNonQuery();

                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = 31;
                        progressBar1.PerformStep();
                        label1.Text = (j - 1) + " of " + 31 + " rows";
                        label1.Refresh();
                    }
                    if (i == xlWorkbook.Sheets.Count)
                    {
                        Console.WriteLine("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                        Proj_LogError("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                    }
                }
            }

        }

        private void retailerSurveyResults(string file)
        {
            server.Filename = file;
            server.Query = "TRUNCATE _retailersurveyresults";
            server.ExecuteNonQuery();
            if (file != "")
            {
                //Read File
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);

                for (var i = 1; i <= xlWorkbook.Sheets.Count; i++)
                {
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[i];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    for (var j = 2; j < 30; j++)//loop through entries aka rows
                    {
                        string queryInsert = "INSERT INTO _retailersurveyresults (";
                        for (var k = 1; k < 82; k++)//get column names for query
                        {
                            if (k == 81) //if it's the last column
                            {
                                queryInsert += "`" + xlRange.Cells[1, k].Value2.Trim() + "`) VALUES (";
                            }
                            else //while it's not the last column
                            {
                                queryInsert += "`" + xlRange.Cells[1, k].Value2.Trim() + "`,";
                            }
                        }
                        for (var m = 1; m < 82; m++)//get values for query
                        {
                            if (m == 81) //if it's the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert += "NULL);";
                                }
                                else
                                {
                                    queryInsert += "'" + xlRange.Cells[j, m].Value2 + "');";
                                }
                                
                            }
                            else //while it's not the last column
                            {
                                if (xlRange.Cells[j, m].Value2 == null)
                                {
                                    queryInsert += "NULL,";
                                }
                                else
                                {
                                    queryInsert += "'" + xlRange.Cells[j, m].Value2 + "',";
                                }
                            }
                        }
                        textBox1.Text = "queryInsert: " + queryInsert;
                        //Console.WriteLine("queryInsert: " + queryInsert);
                        server.Row = j;
                        server.Query = queryInsert;
                        server.ExecuteNonQuery();

                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = 28;
                        progressBar1.PerformStep();
                        label1.Text = (j - 1) + " of " + 28 + " rows";
                        label1.Refresh();
                    }
                    if (i == xlWorkbook.Sheets.Count)
                    {
                        Console.WriteLine("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                        Proj_LogError("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                    }
                }
            }

        }

        private string getTableName(string viewname)
        {
            return viewname.Replace("View", "Vtable");
        }

        private void Proj_LogError(string message)
        {
            StringBuilder bld = new StringBuilder();
            bld.AppendLine(message);
            bld.AppendLine(txtLog.Text);
            txtLog.Text = bld.ToString();
        }

        private void IncreaseProgressBar()
        {
            progressBar1.Increment(1);
        }
    }
}
