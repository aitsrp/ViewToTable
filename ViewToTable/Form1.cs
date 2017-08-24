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
                if (ext == ".xls" || ext == ".xlsx" || ext == ".vls" || ext == ".sql")
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
                    if (inf.Extension == ".vls")
                    {
                        List<string> viewname = new List<string>();
                        var completed = 0;

                        string[] lines = System.IO.File.ReadAllLines(path);
                        foreach (string line in lines)
                        {
                            if (line.StartsWith("//") || (line == ""))
                            {
                                Console.WriteLine("ignored: " + line);
                            }
                            else
                            {
                                if (line.Contains(", "))
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
                            label1.Text = "Processing view " + completed + " of " + viewname.Count + " views";
                            label1.Refresh();
                        }
                    }
                    else if (inf.Extension == ".sql")
                    {
                        string content = File.ReadAllText(path);
                        string[] lst = content.Split(';');
                        foreach (string str in lst)
                        {
                            if (str.Trim() != "")
                            {
                                server.Query = str;
                                server.ExecuteNonQuery();
                            }
                        }

                    }
                    else if (filename.Contains("234584")) //General Survey Results
                    {
                        Console.WriteLine("Processing General Survey Results...");
                        Proj_LogError("Processing General Survey Results...");
                        generalSurveyResults(path, "_generalsurveyresults");
                    }
                    else if (filename.Contains("672569")) //General Survey Results Midpoint
                    {
                        Console.WriteLine("Processing General Survey Results Midpoint...");
                        Proj_LogError("Processing General Survey Results Midpoint...");
                        generalSurveyResults(path, "_generalsurveyresultsm");
                    }
                    else if (filename.Contains("787585")) //General Survey Results Final
                    {
                        Console.WriteLine("Processing General Survey Results Final...");
                        Proj_LogError("Processing General Survey Results Final...");
                        generalSurveyResults(path, "_generalsurveyresultsf");
                    }
                    else if (filename.Contains("972221")) // Product Survey Results
                    {
                        Console.WriteLine("Processing Product Survey Results...");
                        Proj_LogError("Processing Product Survey Results...");
                        otherSurveyResults(path, "_productsurveyresults");
                    }
                    else if (filename.Contains("788185")) // Product Survey Results Final
                    {
                        Console.WriteLine("Processing Product Survey Results Final...");
                        Proj_LogError("Processing Product Survey Results Final...");
                        otherSurveyResults(path, "_productsurveyresultsf");
                    }
                    else if (filename.Contains("977717")) //Retailer Survey Results
                    {
                        Console.WriteLine("Processing Retailer Survey Results...");
                        Proj_LogError("Processing Retailer Survey Results...");
                        otherSurveyResults(path, "_retailersurveyresults");
                    }
                    else if (filename.Contains("818999")) //Retailer Survey Results Midpoint
                    {
                        Console.WriteLine("Processing Retailer Survey Results Midpoint...");
                        Proj_LogError("Processing Retailer Survey Results Midpoint...");
                        otherSurveyResults(path, "_retailersurveyresultsm");
                    }
                    else if (filename.Contains("473321")) //Retailer Survey Results Final
                    {
                        Console.WriteLine("Processing Retailer Survey Results Final...");
                        Proj_LogError("Processing Retailer Survey Results Final...");
                        otherSurveyResults(path, "_retailersurveyresultsf");
                    }
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
            textBox1.Refresh();
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
                textBox1.Refresh();
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
                    textBox1.Refresh();
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

        private void generalSurveyResults(string file, string tablename)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);

            if (file != "")
            {
                List<List<string>> list = SheetToList(xlWorkbook.Sheets[1]);
                var cols = list[0].Count;
                var rows = list.Count;
                progressBar1.Value = 0;
                server.Filename = file;

                string queryTable1 = "SELECT 1 FROM " + tablename + "1 LIMIT 1;";
                server.Query = queryTable1;
                textBox1.Text = "queryView: " + queryTable1;
                textBox1.Refresh();
                //Console.WriteLine("queryTable1: " + queryTable1);
                bool exists1 = server.ExecuteNonQuery();
                if (!exists1)
                {
                    createTable((tablename + "1"), list, 1, 132);
                }

                string queryTable2 = "SELECT 1 FROM " + tablename + "2 LIMIT 1;";
                server.Query = queryTable2;
                textBox1.Text = "queryView: " + queryTable2;
                textBox1.Refresh();
                //Console.WriteLine("queryTable2: " + queryTable2);
                bool exists2 = server.ExecuteNonQuery();
                if (!exists2)
                {
                    createTable((tablename + "2"), list, 132, cols);
                }

                server.Query = "TRUNCATE " + tablename + "1;";
                server.ExecuteNonQuery();
                server.Query = "TRUNCATE " + tablename + "2;";
                server.ExecuteNonQuery();
                server.Query = "SELECT * FROM " + tablename + "1;";
                DataTable dt1 = server.ExecuteQuery();
                server.Query = "SELECT * FROM " + tablename + "2;";
                DataTable dt2 = server.ExecuteQuery();


                for (var j = 1; j < (rows); j++)//loop through entries aka rows
                {
                    //generalsurveyresults1
                    string queryInsert1 = "INSERT INTO " + tablename + "1 (";
                    for (var k = 0; k < 132; k++)//get column names for query
                    {
                        if (k == 131) //if it's the last column
                        {
                            queryInsert1 += "`" + list[0][k] + "`) VALUES (";
                        }
                        else //while it's not the last column
                        {
                            queryInsert1 += "`" + list[0][k] + "`,";
                        }
                    }
                    for (var m = 0; m < 132; m++)//get values for query
                    {
                        if (m == 131) //if it's the last column
                        {
                            if (list[j][m] == null)
                            {
                                queryInsert1 += "NULL);";
                            }
                            else
                            {
                                bool result = dt1.Columns[(m - 1)].DataType == System.Type.GetType("System.String");
                                if (result)
                                {
                                    queryInsert1 += "'" + list[j][m].Replace("'", "\'") + "');";
                                }
                                else
                                    queryInsert1 += "'" + list[j][m] + "');";
                            }
                        }
                        else //while it's not the last column
                        {
                            if (list[j][m] == null)
                            {
                                queryInsert1 += "NULL,";
                            }
                            else
                            {
                                bool result = dt1.Columns[(m)].DataType == System.Type.GetType("System.String");
                                if (result)
                                {
                                    queryInsert1 += "'" + list[j][m].Replace("'", "\'") + "',";
                                }
                                else
                                    queryInsert1 += "'" + list[j][m] + "',";

                            }
                        }
                    }
                    textBox1.Text = "queryInsert1: " + queryInsert1;
                    textBox1.Refresh();
                    //Console.WriteLine("queryInsert1: " + queryInsert1);
                    server.Row = j;
                    server.Query = queryInsert1;
                    server.ExecuteNonQuery();

                    //generalsurveyresults2
                    string queryInsert2 = "INSERT INTO " + tablename + "2 (`" + list[0][0] + "`,";
                    for (var n = 132; n < (cols); n++)//get column names for query
                    {
                        if (n == (cols - 1)) //if it's the last column
                        {
                            queryInsert2 += "`" + list[0][n] + "`) VALUES ('" + list[j][0] + "',";
                        }
                        else //while it's not the last column
                        {
                            queryInsert2 += "`" + list[0][n] + "`,";
                        }
                    }
                    for (var o = 132; o < (cols); o++)//get values for query
                    {
                        if (o == (cols - 1)) //if it's the last column
                        {
                            if (list[j][o] == null)
                            {
                                queryInsert2 += "NULL);";
                            }
                            else
                            {

                                bool result = dt2.Columns[(o - 132)].DataType == System.Type.GetType("System.String");
                                if (result)
                                {
                                    queryInsert2 += "'" + list[j][o].Replace("'", "\'") + "');";
                                }
                                else
                                    queryInsert2 += "'" + list[j][o] + "');";
                            }
                        }
                        else //while it's not the last column
                        {
                            if (list[j][o] == null)
                            {
                                queryInsert2 += "NULL,";
                            }
                            else
                            {
                                //Console.WriteLine(dt2.Columns[(o - 1)].DataType == System.Type.GetType("System.String"));
                                bool result = dt2.Columns[(o - 132)].DataType == System.Type.GetType("System.String");
                                if (result)
                                {
                                    queryInsert2 += "'" + list[j][o].Replace("'", "\'") + "',";
                                }
                                else
                                    queryInsert2 += "'" + list[j][o] + "',";
                            }
                        }
                    }
                    textBox1.Text = "queryInsert2: " + queryInsert2;
                    textBox1.Refresh();
                    //Console.WriteLine("queryInsert2: " + queryInsert2);
                    server.Query = queryInsert2;
                    server.ExecuteNonQuery();

                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = rows - 1;
                    progressBar1.PerformStep();
                    label1.Text = "Processing " + (j) + " of " + (rows - 1) + " rows";
                    label1.Refresh();
                }
            }
            Console.WriteLine("Workbook '" + xlWorkbook.Name + "' done.");
            Proj_LogError("Workbook '" + xlWorkbook.Name + "' done.");
        }

        private void otherSurveyResults(string file, string tablename)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);
            if (file != "")
            {
                List<List<string>> list = SheetToList(xlWorkbook.Sheets[1]);
                var cols = list[0].Count;
                var rows = list.Count;
                progressBar1.Value = 0;
                server.Filename = file;

                string queryTable = "SELECT 1 FROM " + tablename + " LIMIT 1;";
                server.Query = queryTable;
                textBox1.Text = "queryTable: " + queryTable;
                textBox1.Refresh();
                //Console.WriteLine("queryTable: " + queryTable);
                bool exists1 = server.ExecuteNonQuery();
                if (!exists1)
                {
                    createTable(tablename, list, 1, cols);
                }

                server.Query = "TRUNCATE " + tablename;
                server.ExecuteNonQuery();
                for (var j = 1; j < (rows); j++)//loop through entries aka rows starting with row 2
                {
                    string queryInsert = "INSERT INTO " + tablename + " (";
                    for (var k = 0; k < cols; k++)//get column names for query
                    {
                        if (k == (cols - 1)) //if it's the last column
                        {
                            queryInsert += "`" + list[0][k] + "`) VALUES (";
                        }
                        else //while it's not the last column
                        {
                            queryInsert += "`" + list[0][k] + "`,";
                        }
                    }
                    for (var m = 0; m < cols; m++)//get values for query
                    {
                        if (m == (cols - 1)) //if it's the last column
                        {
                            if (list[j][m] == null)
                            {
                                queryInsert += "NULL);";
                            }
                            else
                            {
                                queryInsert += "'" + list[j][m] + "');";
                            }
                        }
                        else //while it's not the last column
                        {
                            if (list[j][m] == null)
                            {
                                queryInsert += "NULL,";
                            }
                            else
                            {
                                queryInsert += "'" + list[j][m] + "',";
                            }

                        }
                    }
                    textBox1.Text = "queryInsert: " + queryInsert;
                    textBox1.Refresh();
                    //Console.WriteLine("queryInsert: " + queryInsert);
                    server.Row = j;
                    server.Query = queryInsert;
                    server.ExecuteNonQuery();
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = list.Count - 1;
                    progressBar1.PerformStep();
                    label1.Text = "Processing " + (j) + " of " + (list.Count - 1) + " rows";
                    label1.Refresh();
                }
            }
            Console.WriteLine("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
            Proj_LogError("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
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
            txtLog.Refresh();
        }

        private List<List<string>> SheetToList(Excel._Worksheet sheet)
        {
            Excel.Range xlRange = sheet.UsedRange;
            var xlRows = xlRange.Rows.Count;
            var xlColumns = xlRange.Columns.Count;
            var xlColumnsA = xlRange.Columns.Address.Replace(xlRows.ToString(), "").Split(':')[1];
            var rowstart = 1;

            List<List<string>> list = new List<List<string>>();
            List<string> item1 = new List<string>();

            for (var i = rowstart; i < (xlRows + 1); i++)
            {
                List<string> item = new List<string>();
                Excel.Range range = sheet.get_Range("A" + i.ToString(), xlColumnsA + i.ToString());
                System.Array itemar = (System.Array)range.Cells.Value;
                for (var j = 1; j < xlColumns + 1; j++)
                {
                    if (itemar.GetValue(1, j) != null)
                    {
                        item.Add((string)itemar.GetValue(1, j).ToString());
                    }
                    else
                        item.Add("");
                }

                list.Add(item);
                var progressbarMax = xlRows - 1;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = progressbarMax;
                progressBar1.PerformStep();
                label1.Text = "Reading " + (i - 1) + " of " + progressbarMax + " rows";
                label1.Refresh();
            }
            Console.WriteLine("Row Count: " + list.Count);

            return list;
        }

        private void createTable (string tablename, List<List<string>> list, int start, int fin)
        {
            string queryCreate = "CREATE TABLE `" + tablename + "` (`" + list[0][0] + "` int(11) DEFAULT NULL,";
            for (var k = start; k < fin; k++)//get column names for query
            {
                if (k == (fin - 1)) //if it's the last column
                {
                    queryCreate += "`" + list[0][k] + "` text) ENGINE = InnoDB DEFAULT CHARSET = utf8;";
                }
                else //while it's not the last column
                {
                    queryCreate += "`" + list[0][k] + "` text,";
                }
            }
            textBox1.Text = "queryCreate: " + queryCreate;
            textBox1.Refresh();
            //Console.WriteLine("queryCreate: " + queryCreate);
            server.Query = queryCreate;
            server.ExecuteNonQuery();
            Proj_LogError("Table '" + tablename + "' created.");
        }
    }
}
