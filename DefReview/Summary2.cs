using DefReview.Properties;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.ComponentModel;

namespace DefReview
{
    internal static class Summary2
    {
        private static List<string> mappMeistrijos;
        private static List<string> mappPavojingumai;
        private static Dictionary<string, AddressTable> dictTables;

        private static void createMappings(string OptionsConnString)
        {
            using (OleDbConnection conn = new OleDbConnection(OptionsConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "SELECT pavojingumas FROM Pavojingumai ORDER BY rikiavimas";
                cmd.Connection = conn;
                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    mappPavojingumai = new List<string>();
                    while (reader.Read())
                    {
                        mappPavojingumai.Add(reader["pavojingumas"].ToString());
                    }
                }
                cmd.CommandText = "SELECT indeksas FROM Meistrijos ORDER BY rikiavimas";
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    mappMeistrijos = new List<string>();
                    while (reader.Read())
                    {
                        mappMeistrijos.Add(reader["indeksas"].ToString());
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
                cmd.CommandText = "SELECT rname, raddress FROM SummaryRanges";
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

        private static int[,] fillTable(string Where, OleDbCommand cmd)
        {
            int[,] arr = new int[mappPavojingumai.Count, mappMeistrijos.Count];
            cmd.CommandText = "SELECT Def_pavoing AS pavojingumas, Meistrija AS meistrijaInd, COUNT(*) AS vnt FROM Defektai " + Where + " GROUP BY Meistrija, Def_pavoing";     
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    int pavojingumasId = mappPavojingumai.IndexOf(reader["pavojingumas"].ToString());
                    int meistrijaId = mappMeistrijos.IndexOf(reader["meistrijaInd"].ToString());
                    int vnt = Convert.ToInt32(reader["vnt"]);
                    arr[pavojingumasId, meistrijaId] = vnt;
                }
            }
            return arr;
        }

        private static void createTables(DateTime from, DateTime to, DateTime bottomDate, string DefectConnString)
        {
            string afterBottomDate = string.Format("Aptik_data >= #{0:d}#", bottomDate);
            string betweenDates = string.Format("BETWEEN #{0:d}# AND #{1:d}#", from, to);
            string afterPeriod = string.Format("> #{0:d}#", to);
            string beforePeriod = string.Format("< #{0:d}#", from);
            string beforeEndOfPeriod = string.Format("<= #{0:d}#", to);
            string outputFolder = Settings.Default.SummaryOutputFolder;

            using (OleDbConnection conn = new OleDbConnection(DefectConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                conn.Open();

                //0 Aptikti
                string Where = "WHERE Aptik_data " + betweenDates;
                dictTables["Aptikta"].IntArray = fillTable(Where, cmd);

                //1 Pakeisti
                Where = "WHERE (" + afterBottomDate + ") AND (FaktPakeistas " + betweenDates + ") AND (Budas = \"pakeista\")";
                dictTables["Pakeista"].IntArray = fillTable(Where, cmd);

                //2 Pervirinti
                Where = "WHERE (" + afterBottomDate + ") AND (FaktPakeistas " + betweenDates + ") AND (Budas = \"pervirinta\")";
                dictTables["Pervirinta"].IntArray = fillTable(Where, cmd);

                //3 Sutvarsliuoti
                Where = "WHERE (" + afterBottomDate + ") AND (FaktPakeistas " + betweenDates + ") AND (Budas = \"sutvarsliuota\")";
                dictTables["Sutvarsliuota"].IntArray = fillTable(Where, cmd);

                //5 nesutvarkyti, termino nėra
                Where = "WHERE (" + afterBottomDate + ") AND ((FaktPakeistas " + afterPeriod + ") OR (FaktPakeistas IS NULL)) AND (TuriButPakeist IS NULL)";
                dictTables["Nepašalinta termino nėra"].IntArray = fillTable(Where, cmd);

                //6 nesutvarkyti, terminas nepasibaigęs
                Where = "WHERE (" + afterBottomDate + ") AND ((FaktPakeistas " + afterPeriod + ") OR (FaktPakeistas IS NULL)) AND (TuriButPakeist IS NOT NULL) AND (TuriButPakeist " + afterPeriod + ")";
                dictTables["Nepašalinta terminas nepasibaigęs"].IntArray = fillTable(Where, cmd);

                //7 nesutvarkyti, terminas pasibaigęs
                Where = "WHERE (" + afterBottomDate + ") AND ((FaktPakeistas " + afterPeriod + ") OR (FaktPakeistas IS NULL)) AND (TuriButPakeist IS NOT NULL) AND (TuriButPakeist " + beforeEndOfPeriod + ")";
                dictTables["Nepašalinta terminas pasibaigęs"].IntArray = fillTable(Where, cmd);
            }
        }

        private static void writeExcel(BackgroundWorker worker, DoWorkEventArgs e, ref int progressCounter, DateTime from, DateTime to)
        {
            string outputFolder = Settings.Default.SummaryOutputFolder;
            string outputFileName = string.Format(Settings.Default.Summary2OutputFileNameTemplate, from, to);
            string fullDestFileName = Path.Combine(outputFolder, outputFileName);
            string templateFileName = Settings.Default.Summary2TemplatePath;
            InteropExcel.Application app = new InteropExcel.Application();
            app.Visible = false;

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(progressCounter++, "Going to copy template file to " + fullDestFileName + "...");

            File.Copy(templateFileName, fullDestFileName, true);

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(progressCounter++, "Template file successfully copied.");

            InteropExcel.Workbook workbook = app.Workbooks.Open(fullDestFileName);
            InteropExcel.Worksheet sheet = (InteropExcel.Worksheet)workbook.Sheets[Settings.Default.Summary2WorksheetName];
            InteropExcel.Range range = null;

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(progressCounter++, "Going change summary heading...");

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(progressCounter++, "Summary heading successfully edited.");

            range = sheet.Range[Settings.Default.Summary2HeaderAddress];
            range.Value = string.Format(range.Value, from, to);

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

        internal static void CreateSummary2(BackgroundWorker worker, DoWorkEventArgs e)
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
                DataChecker.checkData(dateBottom, dateFrom, dateTo, defectConnString, mappMeistrijos, mappPavojingumai);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Checking data error. " + ex.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Defect data is OK.");
            worker.ReportProgress(pc++, "Going to create data tables for work...");
                createTables(dateFrom, dateTo, dateBottom, defectConnString);

            try
            {
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
            MessageBox.Show("Done");

            e.Result = "Success. Enjoy your gorgeous summary.";
        }

    }
}
