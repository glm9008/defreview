using DefReview.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.ComponentModel;

namespace DefReview
{

    public partial class FormMain : Form
    {
        private DateTime bottomDate = new DateTime(2016, 1, 1);

        private string DefectConnString;
        private string OptionsConnString;

        public FormMain()
        {
            InitializeComponent();
            InitializeBackgroundWorker(bgwKmvsQueries, new DoWorkEventHandler(bgwKmvsQueries_DoWork));
            InitializeBackgroundWorker(bgwSummary1, new DoWorkEventHandler(bgwSummary1_DoWork));
            InitializeBackgroundWorker(bgwSummary2, new DoWorkEventHandler(bgwSummary2_DoWork));
        }

        private void InitializeBackgroundWorker(BackgroundWorker bgw, DoWorkEventHandler handler)
        {            
            bgw.WorkerReportsProgress = true;
            bgw.WorkerSupportsCancellation = true;
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.DoWork += handler;
        }

        private void getReadyForBackgroundWork()
        {
            // Reset the text in the result text.
            txbProgress.Text = string.Empty;

            // Disable all buttons.
            btnCreateKmvsQueries.Enabled = false;
            btnCreateSummary2.Enabled = false;
            btnKMVsQueryOutputFolder.Enabled = false;
            btnSummaryOutputFolder.Enabled = false;
            btnKMVsTemplateFile.Enabled = false;
            btnSummary2TemplateFile.Enabled = false;
            btnSetDefectDb.Enabled = false;
            btnSetOptionsDb.Enabled = false;
            btnSaveSettings.Enabled = false;

            // Enable the Cancel button.
            btnCancel.Enabled = true;
        }

        private void returnFromBackgroundWork()
        {
            // Enable all Buttons
            btnCreateKmvsQueries.Enabled = true;
            btnCreateSummary2.Enabled = true;
            btnKMVsQueryOutputFolder.Enabled = true;
            btnSummaryOutputFolder.Enabled = true;
            btnKMVsTemplateFile.Enabled = true;
            btnSummary2TemplateFile.Enabled = true;
            btnSetDefectDb.Enabled = true;
            btnSetOptionsDb.Enabled = true;
            btnSaveSettings.Enabled = true;
            
            // Disable the Cancel button.
            btnCancel.Enabled = false;
        }

        // This event handler updates the progress bar.
        private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (this.txbProgress.Text != string.Empty)
                this.txbProgress.AppendText("\r\n");
            this.txbProgress.AppendText(e.ProgressPercentage.ToString() + ". " + e.UserState.ToString());
        }

        // This event handler deals with the results of the
        // background operation.
        private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                txbProgress.AppendText("\r\nAn error occured.");
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                txbProgress.AppendText("\r\nCanceled.");
            }
            else if (e.Result != null)
            {
                txbProgress.AppendText("\r\n" + e.Result.ToString());
            }
            else
            {
                txbProgress.AppendText("\r\n something went wrong.");
            }
            returnFromBackgroundWork();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // this.bgwKmvsQueries.CancelAsync();
            BackgroundWorker worker = sender as BackgroundWorker;
            worker.CancelAsync();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            setConnectionString();
            setSettings(); 
        }


        private void btnCreateKmvsQueries_Click(object sender, EventArgs e)
        {
            if (bgwKmvsQueries.IsBusy != true)
            {
                Object[] argument = { bottomDate, DefectConnString, OptionsConnString };
                bgwKmvsQueries.RunWorkerAsync(argument);
            }
        }

        private void bgwKmvsQueries_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            QueryKmvs.createKmvsQueries(worker, e);
        }


        private void btnCreateSummary1_Click(object sender, EventArgs e)
        {
            if (bgwSummary1.IsBusy != true)
            {
                Object[] argument = { dtpFrom.Value, dtpTo.Value, bottomDate, DefectConnString, OptionsConnString };
                bgwSummary1.RunWorkerAsync(argument);
            }
        }

        private void bgwSummary1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Summary1.CreateSummary1(worker, e);
        }


        private void btnCreateSummary2_Click(object sender, EventArgs e)
        {
            if (bgwSummary2.IsBusy != true)
            {
                Object[] argument = { dtpFrom.Value, dtpTo.Value, bottomDate, DefectConnString, OptionsConnString };
                bgwSummary2.RunWorkerAsync(argument);
            }
        }

        private void bgwSummary2_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Summary2.CreateSummary2(worker, e);
        }

        private void setConnectionString()
        {
            DefectConnString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;", Settings.Default.DefectDBPath);
            OptionsConnString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;", Settings.Default.OptionsDBPath);
        }
    }
}
