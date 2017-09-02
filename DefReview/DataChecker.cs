using DefReview.Properties;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace DefReview
{
    internal static class DataChecker
    {
        // tikrina, ar duomenų bazėje nėra klaidingų įrašų

        internal static void checkData(DateTime bottom, DateTime from, DateTime to, string DefectConnString, List<string> mappMeistrijos, List<string> mappPavojingumai, List<string> mappKategorijos)
        {
            string betweenDates = string.Format("((Aptik_data BETWEEN #{0:d}# AND #{1:d}#) OR (FaktPakeistas BETWEEN #{0:d}# AND #{1:d}#) OR (Aptik_data >= #{2:d}# AND FaktPakeistas IS NULL))", from, to, bottom);
            string errorFileName = Path.Combine(Properties.Settings.Default.SummaryOutputFolder, string.Format(Settings.Default.Summary2ErrorFileNameTemplate, from, to));

            using (OleDbConnection conn = new OleDbConnection(DefectConnString))
            {
                checkPavojingumai(betweenDates, conn, errorFileName, mappPavojingumai);
                checkMeistrijos(betweenDates, conn, errorFileName, mappMeistrijos);
                checkKategorijos(betweenDates, conn, errorFileName);
                checkBudas(betweenDates, conn, errorFileName);
                checkRadimoLaterAtlikimo(betweenDates, conn, errorFileName);
                checkRadimoLaterTerminas(betweenDates, conn, errorFileName);
            }
        }

        internal static void checkData(DateTime bottom, DateTime from, DateTime to, string DefectConnString, List<string> mappMeistrijos, List<string> mappPavojingumai)
        {
            string betweenDates = string.Format("((Aptik_data BETWEEN #{0:d}# AND #{1:d}#) OR (FaktPakeistas BETWEEN #{0:d}# AND #{1:d}#) OR (Aptik_data >= #{2:d}# AND FaktPakeistas IS NULL))", from, to, bottom);
            string errorFileName = Path.Combine(Properties.Settings.Default.SummaryOutputFolder, string.Format(Settings.Default.Summary2ErrorFileNameTemplate, from, to));

            using (OleDbConnection conn = new OleDbConnection(DefectConnString))
            {
                checkPavojingumai(betweenDates, conn, errorFileName, mappPavojingumai);
                checkMeistrijos(betweenDates, conn, errorFileName, mappMeistrijos);
                checkBudas(betweenDates, conn, errorFileName);
                checkRadimoLaterAtlikimo(betweenDates, conn, errorFileName);
                checkRadimoLaterTerminas(betweenDates, conn, errorFileName);
            }
        }        

        private static void checkRadimoLaterTerminas(string betweenDates, OleDbConnection conn, string errorFileName)
        {
            // patikrina, ar nėra įrašu, kad radimo data vėlesnė negu atlikimo data
            string commandText = "SELECT skait, Aptik_data, TuriButPakeist FROM Defektai WHERE " + betweenDates + " AND (TuriButPakeist IS NOT NULL) AND (Aptik_data > TuriButPakeist)";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait\tradimo data\tsutvarkymo terminas");
                while (reader.Read())
                {
                    sb.AppendLine(string.Format("{0}\t{1:d}\t{2:d}", reader["skait"], reader["Aptik_data"], reader["TuriButPakeist"]));
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra defektų, kurių radimo data yra vėlesnė nei galutinė sutvarkymo data, žr. \"" + errorFileName + "\"");
            }
        }

        private static void checkRadimoLaterAtlikimo(string betweenDates, OleDbConnection conn, string errorFileName)
        {
            // patikrina, ar nėra įrašu, kad radimo data vėlesnė negu atlikimo data
            string commandText = "SELECT skait, Aptik_data, FaktPakeistas FROM Defektai WHERE " + betweenDates + " AND (FaktPakeistas IS NOT NULL) AND (Aptik_data > FaktPakeistas)";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait\tradimo data\tatlikimo data");
                while (reader.Read())
                {
                    sb.AppendLine(string.Format("{0}\t{1:d}\t{2:d}", reader["skait"], reader["Aptik_data"], reader["FaktPakeistas"]));
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra defektų, kurių radimo data yra vėlesnė nei sutvarkymo data, žr. \"" + errorFileName + "\"");
            }
        }

        private static void checkBudas(string betweenDates, OleDbConnection conn, string errorFileName)
        {
            // patikrina, ar nėra įrašu, kad atlikta, bet nenurodytas būdas
            string commandText = "SELECT skait FROM Defektai WHERE " + betweenDates + " AND (FaktPakeistas IS NOT NULL) AND (Budas IS NULL OR Budas = \"\")";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait");
                while (reader.Read())
                {
                    sb.AppendLine(reader["skait"].ToString());
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra defektų, kurie atlikti, bet nenurodytas atlikimo būdas, žr. \"" + errorFileName + "\"");
            }
        }

        private static void checkMeistrijos(string betweenDates, OleDbConnection conn, string errorFileName, List<string> mappMeistrijos)
        {
            // patikrina ar yra meistrijų, kurios nebūtų mappintos arba ar yra meistrijų, kurios yra null
            string meiMap = string.Join(",", mappMeistrijos.Select(mei => "\"" + mei + "\"").ToArray());
            string commandText = "SELECT skait, Meistrija FROM Defektai WHERE " + betweenDates + " AND ((NOT Meistrija IN (" + meiMap + ")) OR (Meistrija IS NULL))";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait\tmeistrija");
                while (reader.Read())
                {
                    sb.AppendLine(reader["skait"] + "\t" + reader["Meistrija"]);
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra klaidų duomenų bazėje - neteisingai įvestos meistrijos, žr. \"" + errorFileName + "\"");
            }
        }

        private static void checkPavojingumai(string betweenDates, OleDbConnection conn, string errorFileName, List<string> mappPavojingumai)
        {
            // patikrina ar yra pavojingumų, kurie nebūtų mappinti arba ar yra pavojingumų, kurie yra null
            string pavMap = string.Join(",", mappPavojingumai.Select(pav => "\"" + pav + "\"").ToArray());
            string commandText = "SELECT skait, Def_pavoing FROM Defektai WHERE " + betweenDates + " AND ((NOT Def_pavoing IN (" + pavMap + ")) OR (Def_pavoing IS NULL))";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait\tpavojingumas");
                while (reader.Read())
                {
                    sb.AppendLine(reader["skait"] + "\t" + reader["Def_pavoing"]);
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra klaidų duomenų bazėje - neteisingai įvesti pavojingumai, žr. \"" + errorFileName + "\"");
            }
        }

        private static void checkKategorijos(string betweenDates, OleDbConnection conn, string errorFileName)
        {
            // patikrina ar yra kategorijų, kurios būtų outside [1; 6] arba yra null
            string commandText = "SELECT skait, LK as kategorija FROM Defektai WHERE " + betweenDates + " AND ((NOT LK BETWEEN 1 AND 6) OR (LK IS NULL))";
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            conn.Open();
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                if (!reader.HasRows)
                {
                    reader.Close();
                    conn.Close();
                    return;
                }

                StringBuilder sb = new StringBuilder().AppendLine("skait\tkategorija");
                while (reader.Read())
                {
                    sb.AppendLine(reader["skait"] + "\t" + reader["kategorija"]);
                }

                writeErrors(errorFileName, sb.ToString());
                reader.Close();
                conn.Close();
                throw new Exception("Yra klaidų duomenų bazėje - neteisingai įvestos kelių kategorijos (LK), žr. \"" + errorFileName + "\"");
            }
        }

        private static void writeErrors(string errorFileName, string errors)
        {
            try
            {
                File.WriteAllText(errorFileName, errors);
            }
            catch (Exception ex)
            {
                throw new Exception("Duomenų bazės įrašuose yra klaidų, bet nepavyksta įrašyti tų klaidų failo.\n" + ex.Message + "\n" + errors);
            }
        }

    }
}
