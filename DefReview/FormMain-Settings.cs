using DefReview.Properties;
using System;
using System.IO;
using System.Windows.Forms;

namespace DefReview
{
    public partial class FormMain
    {
        private void setSettings()
        {
            // Common Settings
            txbDefectDbPath.Text = Settings.Default.DefectDBPath;
            txbOptionsDbPath.Text = Settings.Default.OptionsDBPath;

            // KMVs Query Settings
            txbTopLeftCell.Text = Settings.Default.KMVsQueryTopLeftCell;
            txbWorkSheetName.Text = Settings.Default.KMVsQueryWorksheetName;
            txbKMVsQueryTemplatePath.Text = Settings.Default.KMVsQueryTemplatePath;
            txbKMVsQueryFileNameTemplate.Text = Settings.Default.KMVsQueryFileNameTemplate;
            txbKMVsQueryOutputFolder.Text = Settings.Default.KMVsQueryOutputFolder;

            // Summary1 Settings
            txbSummary1OutputFileNameTemplate.Text = Settings.Default.Summary1OutputFileNameTemplate;
            txbSummary1ErrorFileNameTemplate.Text = Settings.Default.Summary1ErrorFileNameTemplate;
            txbSummary1HeaderAddress.Text = Settings.Default.Summary1HeaderAddress;
            txbSummary1TemplatePath.Text = Settings.Default.Summary1TemplatePath;
            txbSummary1WorksheetName.Text = Settings.Default.Summary1WorksheetName;

            // Summary2 Settings
            txbSummary2OutputFileNameTemplate.Text = Settings.Default.Summary2OutputFileNameTemplate;
            txbSummary2ErrorFileNameTemplate.Text = Settings.Default.Summary2ErrorFileNameTemplate;
            txbSummary2HeaderAddress.Text = Settings.Default.Summary2HeaderAddress;
            txbSummary2TemplatePath.Text = Settings.Default.Summary2TemplatePath;
            txbSummary2WorksheetName.Text = Settings.Default.Summary2WorksheetName;

            // Summary common settings
            txbSummaryOutputFolder.Text = Settings.Default.SummaryOutputFolder;
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            if (
                    CheckTextBox(txbDefectDbPath, "DefectDBPath") &&
                    CheckTextBox(txbOptionsDbPath, "OptionsDBPath") &&
                    CheckTextBox(txbTopLeftCell, "KMVsQueryTopLeftCell") &&
                    CheckTextBox(txbWorkSheetName, "KMVsQueryWorksheetName") &&
                    CheckTextBox(txbKMVsQueryTemplatePath, "KMVsQueryTemplatePath") &&
                    CheckTextBox(txbKMVsQueryFileNameTemplate, "KMVsQueryFileNameTemplate") &&
                    CheckTextBox(txbKMVsQueryOutputFolder, "KMVsQueryOutputFolder") &&
                    CheckTextBox(txbSummary2WorksheetName, "Summary2WorksheetName") &&
                    CheckTextBox(txbSummary1OutputFileNameTemplate, "Summary1OutputFileNameTemplate") &&
                    CheckTextBox(txbSummary1ErrorFileNameTemplate, "Summary1ErrorFileNameTemplate") &&
                    CheckTextBox(txbSummary1HeaderAddress, "Summary1HeaderAddress") &&
                    CheckTextBox(txbSummary1TemplatePath, "Summary1TemplatePath") &&
                    CheckTextBox(txbSummary1WorksheetName, "Summary1WorksheetName") &&
                    CheckTextBox(txbSummary2OutputFileNameTemplate, "Summary2OutputFileNameTemplate") &&
                    CheckTextBox(txbSummary2ErrorFileNameTemplate, "Summary2ErrorFileNameTemplate") &&
                    CheckTextBox(txbSummary2HeaderAddress, "Summary2HeaderAddress") &&
                    CheckTextBox(txbSummary2TemplatePath, "Summary2TemplatePath") &&
                    CheckTextBox(txbSummaryOutputFolder, "SummaryOutputFolder")
                )
            {
                Settings.Default.Save();
                MessageBox.Show("Nustatymai pakeisti", "Atlikta gerai", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            else 
            {
                MessageBox.Show("Teksto laukai negali būti tušti.", "Neteisingai užpildyta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private bool CheckTextBox(TextBox txb, string setting)
        {
            string text = txb.Text.Trim();
            if (text.Length > 0)
            {
                Settings.Default[setting] = text;
                return true;
            }
            else
            {
                
                return false;
            }
        }

        private void setFile(string pathSetting, TextBox txb, string title, string filter)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Title = title;
            if (Settings.Default[pathSetting] == null || !File.Exists((string)(Settings.Default[pathSetting])))
            {
                ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            else
            {
                ofd.InitialDirectory = Path.GetDirectoryName((string)(Settings.Default[pathSetting]));
            }
            ofd.Filter = filter;
            ofd.FilterIndex = 1;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Settings.Default[pathSetting] = ofd.FileName;
                txb.Text = ofd.FileName;
                Settings.Default.Save();
            }
            ofd.Dispose();
        }

        private void btnSetDefectDb_Click(object sender, EventArgs e)
        {
            setFile("DefectDBPath",  
                txbDefectDbPath, 
                "Defektų duomenų bazė", 
                "Access (*.accdb;*.mdb)|*.accdb;*.mdb|All files (*.*)|*.*");
            setConnectionString();
        }

        private void btnSetOptionsDb_Click(object sender, EventArgs e)
        {
            setFile("OptionsDBPath", 
                txbOptionsDbPath,
                "Pagalbinė duomenų bazė",
                "Access (*.accdb;*.mdb)|*.accdb;*.mdb|All files (*.*)|*.*");
            setConnectionString();
        }

        private void btnKMVsTemplateFile_Click(object sender, EventArgs e)
        {
            setFile("KMVsQueryTemplatePath", 
                txbKMVsQueryTemplatePath, 
                "KMVs Query Template File", 
                "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*");
        }

        private void btnSummary1TemplateFile_Click(object sender, EventArgs e)
        {
            setFile("Summary1TemplatePath",
                txbSummary1TemplatePath,
                "Summary 1 Template File",
                "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*");
        }

        private void btnSummary2TemplateFile_Click(object sender, EventArgs e)
        {
            setFile("Summary2TemplatePath",
                txbSummary2TemplatePath,
                "Summary 2 Template File",
                "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*");
        }

        private void setFolder(string ofSetting, TextBox txb)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = true;
            if (Settings.Default[ofSetting] == null || !Directory.Exists((string)(Settings.Default[ofSetting])))
            {
                fbd.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                fbd.SelectedPath = (string)(Settings.Default[ofSetting]);
            }
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                Settings.Default[ofSetting] = fbd.SelectedPath;
                txb.Text = fbd.SelectedPath;
                Settings.Default.Save();
            }
        }

        private void btnKMVsQueryOutputFolder_Click(object sender, EventArgs e)
        {
            setFolder("KMVsQueryOutputFolder", txbKMVsQueryOutputFolder);
        }

        private void btnSummaryOutputFolder_Click(object sender, EventArgs e)
        {
            setFolder("SummaryOutputFolder", txbSummaryOutputFolder);
        }        
    }
}
