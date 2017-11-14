using DefReview.Properties;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.ComponentModel;

namespace DefReview
{

    public class AddressTable
    {
        public string Address { get; set; }
        public int[,] IntArray { get; set; }
        public AddressTable(string address)
        {
            this.Address = address;
        }
    }
	
    internal static class Summary1
    {
        
        private static List<string> mappPavojingumai;
        private static List<string> mappMeistrijos;
        private static List<string> mappKategorijos;
        private static Dictionary<string, AddressTable> dictTables;

        private static void createMappings(string OptionsConnString)
        {
            using (OleDbConnection conn = new OleDbConnection(OptionsConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "SELECT indeksas FROM Meistrijos ORDER BY rikiavimas";
                cmd.Connection = conn;
                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    mappMeistrijos = new List<string>();
                    while (reader.Read())
                    {
                       mappMeistrijos.Add(reader["indeksas"].ToString());
                    }
                }
                cmd.CommandText = "SELECT pavojingumas FROM Pavojingumai ORDER BY rikiavimas";
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    mappPavojingumai = new List<string>();
                    while (reader.Read())
                    {
                        mappPavojingumai.Add(reader["pavojingumas"].ToString());
                    }
                }
                cmd.CommandText = "SELECT kategorija FROM Kategorijos ORDER BY rikiavimas";
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    mappKategorijos = new List<string>();
                    while (reader.Read())
                    {
                        mappKategorijos.Add(reader["kategorija"].ToString());
                    }
                }
            }
        }

        private static void loadRanges(string OptionsConnString)
        {
            dictTables = new Dictionary<string, AddressTable>();
            using (OleDbConnection conn = new OleDbConnection(OptionsConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "SELECT rname, raddress FROM Summary1Ranges";
                cmd.Connection = conn;
                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        dictTables.Add(reader["rname"].ToString(), new AddressTable(reader["raddress"].ToString()));
                    }
                }
            }
        }

        private static int[,] fillArray(string Where, OleDbCommand cmd)
        {
            int[,] arr = new int[mappMeistrijos.Count, 1];
            cmd.CommandText = "SELECT Meistrija AS meistrijaInd, Count(*) AS vnt FROM Defektai " + Where + " GROUP BY Meistrija";
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    int meistrijaId = mappMeistrijos.IndexOf(reader["meistrijaInd"].ToString());
                    int vnt = Convert.ToInt32(reader["vnt"]);
                    arr[meistrijaId, 0] = vnt;
                }
            }
            return arr;
        }

        private static int[,] fillTable(string Where, OleDbCommand cmd)
        {
            int[,] arr = new int[mappMeistrijos.Count, mappKategorijos.Count];
            cmd.CommandText = "SELECT Meistrija AS meistrijaInd, LK AS kategorija, Count(*) AS vnt FROM Defektai " + Where + " GROUP BY Meistrija, LK";
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string kategorija = reader["kategorija"].ToString();
                    int meistrijaId = mappMeistrijos.IndexOf(reader["meistrijaInd"].ToString());
                    int kategorijaId = -1;
                    switch (reader["kategorija"].ToString())
                    {
                        case "1":
                            kategorijaId = 0;
                            break;
                        case "2":
                            kategorijaId = 1;
                            break;
                        case "3":
                        case "4":
                            kategorijaId = 2;
                            break;
                        default:
                            kategorijaId = 3;
                            break;
                    }
                    int vnt = Convert.ToInt32(reader["vnt"]);
                    arr[meistrijaId, kategorijaId] += vnt;
                }
            }
            return arr;
        }


        private static void createTables(DateTime from, DateTime to, DateTime bottomDate, string DefectConnString)
        {
            string aptiktaFromBottomUptoEndOfPeriod = string.Format("Aptik_data BETWEEN #{0:d}# AND #{1:d}#", bottomDate, to);
            string aptiktaBetweenFromAndTo =  string.Format("Aptik_data BETWEEN #{0:d}# AND #{1:d}#", from, to);
            string sutvarkytaBetweenFromAndTo = string.Format("Faktpakeistas BETWEEN #{0:d}# AND #{1:d}#", from, to);
            string nesutvarkyta = "FaktPakeistas IS NULL";
            string sutvarkytaAfterPeriod = string.Format("Faktpakeistas > #{0:d}#", to);
            string outputFolder = Settings.Default.SummaryOutputFolder;
            using (OleDbConnection conn = new OleDbConnection(DefectConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                conn.Open();
                // Aptikti visi
                string Where = "WHERE " + aptiktaBetweenFromAndTo;
                int[,] aptiktiVisi = fillArray(Where, cmd);

                // Aptikta stebimų
                // Reikia išrinkti: aptik_data - per periodą && (D1 - 5-6 kategorija || D2 - 2-6 kategorija || D3)
                Where = "WHERE (" + aptiktaBetweenFromAndTo + ") AND (((Def_pavoing = \"D1\") AND (LK BETWEEN 5 AND 6)) OR ((Def_pavoing = \"D2\") AND (LK BETWEEN 2 AND 6)) OR (Def_pavoing = \"D3\"))";
                dictTables["Aptikta stebimų"].IntArray = fillArray(Where, cmd);

                // Aptikta ID
                // Reikia išrinkti: aptik_data - per periodą && (D1 - 5-6 kategorija || D2 - 2-6 kategorija || D3)
                Where = "WHERE (" + aptiktaBetweenFromAndTo + ") AND (Def_pavoing = \"ID\")";
                dictTables["Aptikta ID"].IntArray = fillArray(Where, cmd);

                dictTables["Aptikta keičiamų"].IntArray = new int[mappMeistrijos.Count, 1];
                for (int i = 0; i < mappMeistrijos.Count; i++)
                {
                    dictTables["Aptikta keičiamų"].IntArray[i, 0] = aptiktiVisi[i, 0] - dictTables["Aptikta stebimų"].IntArray[i, 0] - dictTables["Aptikta ID"].IntArray[i, 0];
                }

                // Pakeista
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND (" + sutvarkytaBetweenFromAndTo + ") AND (Budas = \"pakeista\")";
                dictTables["Pakeista"].IntArray = fillArray(Where, cmd);

                // Sutvarsliuoti
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND (" + sutvarkytaBetweenFromAndTo + ") AND (Budas = \"sutvarsliuota\")";
                dictTables["Sutvarsliuota"].IntArray = fillArray(Where, cmd);

                // Pervirinti
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND (" + sutvarkytaBetweenFromAndTo + ") AND (Budas = \"pervirinta\")";
                dictTables["Pervirinta"].IntArray = fillArray(Where, cmd);

                
                // Liko DP
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND ((" + sutvarkytaAfterPeriod + ") OR (" + nesutvarkyta + ")) AND (Def_pavoing = \"DP\")";
                dictTables["Liko DP"].IntArray = fillTable(Where, cmd);

                // Liko D1
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND ((" + sutvarkytaAfterPeriod + ") OR (" + nesutvarkyta + ")) AND (Def_pavoing = \"D1\")";
                dictTables["Liko D1"].IntArray = fillTable(Where, cmd);

                // Liko D2
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND ((" + sutvarkytaAfterPeriod + ") OR (" + nesutvarkyta + ")) AND (Def_pavoing = \"D2\")";
                dictTables["Liko D2"].IntArray = fillTable(Where, cmd);

                // Liko D3
                Where = "WHERE (" + aptiktaFromBottomUptoEndOfPeriod + ") AND ((" + sutvarkytaAfterPeriod + ") OR (" + nesutvarkyta + ")) AND (Def_pavoing = \"D3\")";
                dictTables["Liko D3"].IntArray = fillTable(Where, cmd);
            }
        }

        private static void writeExcel(BackgroundWorker worker, DoWorkEventArgs e, ref int progressCounter, DateTime from, DateTime to)
        {
                string outputFolder = Settings.Default.SummaryOutputFolder;
                string outputFileName = string.Format(Settings.Default.Summary1OutputFileNameTemplate, from, to);
                string fullDestFileName = Path.Combine(outputFolder, outputFileName);
                string templateFileName = Settings.Default.Summary1TemplatePath;
                InteropExcel.Application app = new InteropExcel.Application();
                app.Visible = false;

                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "Going to copy template file to " + fullDestFileName + "...");

                File.Copy(templateFileName, fullDestFileName, true);

                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "Template file successfully copied.");

                InteropExcel.Workbook workbook = app.Workbooks.Open(fullDestFileName);
                InteropExcel.Worksheet sheet = (InteropExcel.Worksheet)workbook.Sheets[Settings.Default.Summary1WorksheetName];
                InteropExcel.Range range = null;

            try
            {
                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "Going change summary heading...");

                range = sheet.Range[Settings.Default.Summary1HeaderAddress];
                range.Value = string.Format(range.Value, from, to);

                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "Summary heading successfully edited.");

                foreach (AddressTable atable in dictTables.Values)
                {
                    if (worker.CancellationPending) e.Cancel = true;
                    worker.ReportProgress(progressCounter++, "Going to load region \"" + atable.Address + "\" with data...");

                    range = sheet.Range[atable.Address].Resize[atable.IntArray.GetLength(0), atable.IntArray.GetLength(1)];
                    range.Value = atable.IntArray;

                    if (worker.CancellationPending) e.Cancel = true;
                    worker.ReportProgress(progressCounter++, "Region \"" + atable.Address + "\" successfully loaded with data.");
                }

                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "All data loaded. Going to save file and quit Interop.Excel...");

                workbook.Save();
                app.Quit();

                if (worker.CancellationPending) e.Cancel = true;
                worker.ReportProgress(progressCounter++, "File has been saved, Interop.Excel successfully quitted.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //range = null;
                //sheet = null;
                //workbook = null;
                //app = null;
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);
            }
        }
        
        internal static void CreateSummary1(BackgroundWorker worker, DoWorkEventArgs e)
        {
            Object[] argument = e.Argument as Object[];
            DateTime dateFrom = Convert.ToDateTime(argument[0]);
            DateTime dateTo = Convert.ToDateTime(argument[1]);
            DateTime dateBottom = Convert.ToDateTime(argument[2]);
            string defectConnString = argument[3].ToString();
            string optionsConnString = argument[4].ToString();


            int pc = 0; // progress count
            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Going to fetch all ranges from the database...");

            try
            {
                loadRanges(optionsConnString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loading ranges error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Ranges have been successfully fetched.");
            worker.ReportProgress(pc++, "Going to create mappings...");

            try
            {
                createMappings(optionsConnString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Create mappings error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Mappings have been successfully created.");
            worker.ReportProgress(pc++, "Going to check defect data consistency...");

            try
            {
                DataChecker.checkData(dateBottom, dateFrom, dateTo, defectConnString, mappMeistrijos, mappPavojingumai, mappKategorijos);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Checking data error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Defect data is OK.");
            worker.ReportProgress(pc++, "Going to create data tables for work...");

            try
            {
                createTables(dateFrom, dateTo, dateBottom, defectConnString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Creating tables error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Data tables have been successfully created.");
            worker.ReportProgress(pc++, "Going to copy XLSX template and load it with data...");

            try
            {
                writeExcel(worker, e, ref pc, dateFrom, dateTo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Writting excel error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            e.Result = "Success. Enjoy your gorgeous summary.";
        }

    }
}
