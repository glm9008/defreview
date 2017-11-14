using DefReview.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.ComponentModel;

namespace DefReview
{       
    public struct NamedTable
    {
        public string name;
        public DataTable table;

        public NamedTable(string tableName, DataTable dataTable)
        {
            name = tableName;
            table = dataTable;
        }
    }

    internal static class QueryKmvs
    {

        internal static void createKmvsQueries(BackgroundWorker worker, DoWorkEventArgs e)
        {
            Object[] argument = e.Argument as Object[];
            DateTime bottomDate = Convert.ToDateTime(argument[0]);
            string defectConnString = argument[1].ToString();
            string optionsConnString = argument[2].ToString();

            int pc = 1; // progress count
            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Going to fetch all undone defects from the database...");

            Dictionary<string, System.Data.DataTable> dict = null;
            try
            {
                System.Data.DataTable dt = fetchNesutvarkyti(bottomDate, defectConnString);
                worker.ReportProgress(pc++, "The undone defects have been fetched.");
                worker.ReportProgress(pc++, "Sorting the undone defects by meistrija and creating separate datatable for each meistrija.");
                dict = createDictionary(dt, optionsConnString);
                worker.ReportProgress(pc++, "The defects have been distributed by meistrijos.");
            }
            catch (Exception ex)
            {
                worker.ReportProgress(pc++, "Something went wrong...");
                MessageBox.Show("Klaida, atliekant DB užklausas. " + ex.Message, "DB klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Creating Interop object Excel.Application.");

            string outputFolder = Properties.Settings.Default.KMVsQueryOutputFolder;
            string templateFileName = Settings.Default.KMVsQueryTemplatePath;
            string worksheetName = Settings.Default.KMVsQueryWorksheetName;
            string startCell = Settings.Default.KMVsQueryTopLeftCell;
            int created = 0;
            StringBuilder sbErrors = new StringBuilder();
            InteropExcel.Application app = new InteropExcel.Application();
            app.Visible = false;

            if (worker.CancellationPending) e.Cancel = true;
            worker.ReportProgress(pc++, "Excel.Application has been created and opened.");
            worker.ReportProgress(pc++, "Going to create an XLSX file for each meistrija.");

            foreach (KeyValuePair<string, System.Data.DataTable> pair in dict)
            {
                InteropExcel.Workbook workbook = null;
                InteropExcel.Worksheet sheet = null;
                InteropExcel.Range range = null;
                string destFileName = Path.Combine(outputFolder, pair.Key + ".xlsx");
                worker.ReportProgress(pc++, "Going to create a file " + destFileName + ".");
                try
                {
                    File.Copy(templateFileName, destFileName, true);
                    worker.ReportProgress(pc++, "A new file " + destFileName + " has been created.");
                    System.Data.DataTable table = pair.Value;
                    workbook = app.Workbooks.Open(destFileName);
                    string[,] valueArray = dataTableToStringArray(table);
                    sheet = (InteropExcel.Worksheet)workbook.Sheets[worksheetName];
                    range = sheet.Range[startCell].Resize[table.Rows.Count, table.Columns.Count];
                    range.Value = valueArray;
                    workbook.Save();
                    worker.ReportProgress(pc++, "The file " + destFileName + " has been loaded with data and saved.");
                    created++;
                }
                catch (Exception ex)
                {
                    worker.ReportProgress(pc++, "An error occured by creating the file " + destFileName + ". Attempting to delete it.");
                    sbErrors.AppendFormat("creating {0} failed: {1}", pair.Key, ex.Message).AppendLine();
                    if (File.Exists(destFileName))
                    {
                        try
                        {
                            File.Delete(destFileName);
                            worker.ReportProgress(pc++, "The file " + destFileName + " has been successfully deleted. Attempting to close its Interop.Workbook object.");
                        }
                        catch
                        {
                            MessageBox.Show(
                                string.Format("Gaminant failą {0}, įvyko klaida. Todėl buvo mėginama failą ištrinti, bet tai nepavyko taip pat. Failas {0} gali būti su klaidingais duomenimis.", pair.Key),
                                "Nepavyko ištrinti klaidingo failo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                }
                finally
                {
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    if (workbook != null)
                    {
                        workbook.Close();
                        Marshal.ReleaseComObject(workbook);
                        worker.ReportProgress(pc++, "The Interop.Workbook object has been closed.");
                    }
                }
            }
            worker.ReportProgress(pc++, "Attempting to quit the Interop.Excel.");
            app.Quit();
            Marshal.ReleaseComObject(app);
            if (worker.CancellationPending) e.Cancel = true;
            // report result
            StringBuilder sbReport = new StringBuilder();
            sbReport.AppendFormat("{0:d} files must have been created.", dict.Count).AppendLine();
            sbReport.AppendFormat("{0:d} ones have been created.", created);
            if (sbErrors.Length > 0)
            {
                sbReport.AppendLine("Errors occured:").Append(sbErrors.ToString());
                e.Result = string.Format("{0:d} files have been created. Some errors occured...", created);
            }
            else
            {
                e.Result = string.Format("Success. {0:d} files have been created.", created);
            }
            MessageBox.Show(sbReport.ToString());
        }

        private static string[,] dataTableToStringArray(System.Data.DataTable dataTable)
        {
            string[,] stringarray = new string[dataTable.Rows.Count, dataTable.Columns.Count];
            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    if (dataTable.Columns[c].DataType == typeof(DateTime))
                    {
                        DateTime date = Convert.ToDateTime(dataTable.Rows[r][c]);
                        if (date > DateTime.MinValue)
                            stringarray[r, c] = Convert.ToDateTime(dataTable.Rows[r][c]).ToString("yyyy-MM-dd");
                        else
                            stringarray[r, c] = "";
                    }
                    else
                    {
                        stringarray[r, c] = dataTable.Rows[r][c].ToString();
                    }                    
                }
            }
            return stringarray;
        }

        private static Dictionary<string, System.Data.DataTable> createDictionary(System.Data.DataTable tblAllRecords, string OptionsConnString)
        {
            Dictionary<string, System.Data.DataTable> dict = new Dictionary<string, System.Data.DataTable>();
            DataView view = new DataView(tblAllRecords);
            using (OleDbConnection conn = new OleDbConnection(OptionsConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "SELECT indeksas, pavadinimas FROM Meistrijos";
                cmd.Connection = conn;
                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string indeksas = reader["indeksas"].ToString();
                        view.RowFilter = "meistrija = '" + indeksas + "'";
                        if (view.Count > 0)
                        {
                            view.Sort = "vkodas";
                            System.Data.DataTable tblCurrent = view.ToTable();
                            tblCurrent.Columns.Remove("meistrija");
                            dict.Add(reader["pavadinimas"].ToString(), tblCurrent);
                        }
                    }
                }
            }
            return dict;
        }

        private static System.Data.DataTable fetchNesutvarkyti(DateTime bottomDate, string DefectConnString)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("id", typeof(Int64));
            dt.Columns.Add("vkodas", typeof(string));
            dt.Columns.Add("defkodas", typeof(string));
            dt.Columns.Add("dpavojing", typeof(string));
            dt.Columns.Add("aptikdata", typeof(DateTime));
            dt.Columns.Add("pakeistiiki", typeof(DateTime));
            dt.Columns.Add("pastaba", typeof(string));
            dt.Columns.Add("meistrija", typeof(string));
            using (OleDbConnection conn = new OleDbConnection(DefectConnString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = string.Format("SELECT Skait, Linija, Kelias, Km, Piketas, Metrai, Siule, Defekto_kodas, Def_pavoing, Aptik_data, TuriButPakeist, Pastaba, Meistrija FROM Defektai WHERE Aptik_data >= #{0:d}# AND FaktPakeistas IS NULL ORDER BY Meistrija, Linija, Kelias, Km, Piketas, Metrai, Siule", bottomDate);
                cmd.Connection = conn;
                conn.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        long id = Convert.ToInt64(reader["Skait"]);
                        string vkodas = vietosKodas(reader["Linija"], reader["Kelias"], reader["Km"], reader["Piketas"], reader["Metrai"], reader.IsDBNull(reader.GetOrdinal("Siule")) ? null : reader["Siule"]);
                        string defkodas = reader["Defekto_kodas"].ToString();
                        string dpavojing = reader["Def_pavoing"].ToString();
                        DateTime aptikdata = reader.IsDBNull(reader.GetOrdinal("Aptik_data")) ? DateTime.MinValue : Convert.ToDateTime(reader["Aptik_data"]);
                        DateTime pakeistiiki = reader.IsDBNull(reader.GetOrdinal("TuriButPakeist")) ? DateTime.MinValue : Convert.ToDateTime(reader["TuriButPakeist"]);
                        string pastaba = reader.IsDBNull(reader.GetOrdinal("Pastaba")) ? "" : reader["Pastaba"].ToString();
                        string meistrija = reader["Meistrija"].ToString();

                        dt.Rows.Add(id, vkodas, defkodas, dpavojing, aptikdata, pakeistiiki, pastaba, meistrija);
                    }
                }
            }

            return dt;
        }

        private static string vietosKodas (object linija, object kelias, object km, object pk, object m, object siule = null)
        {
            if (siule == null)
            {
                return string.Format("{0}.{1:0}{2:000}.{3:00}.{4:00}", linija, kelias, km, pk, m);
            }
            else
            {
                return string.Format("{0}.{1:0}{2:000}.{3:00}.{4:00}.{5:0}", linija, kelias, km, pk, m, siule);
            }
        }

    }
}
