using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using B1PostingDataText.B1Connection;
using B1PostingDataText.Functions;
using System.Web.Script.Serialization;

namespace B1PostingDataText
{
    public partial class FormTransServices : Form
    {
        public FormTransServices(bool buttonClick = false)
        {
            InitializeComponent();
            ButtonClick = buttonClick;
        }
        public bool ButtonClick { get; set; }

        public string TxtIsrunning { get { return Application.StartupPath+@"\TransIsRunning.txt"; } }
        private void FormTransServices_Load(object sender, EventArgs e)
        {
            if(ButtonClick)
            {
                this.WindowState = FormWindowState.Normal;
            }
            if (!File.Exists(TxtIsrunning))
            {
                TransactionBWorker.ProgressChanged += new ProgressChangedEventHandler(TransactionBWorker_ProgressChanged);
                TransactionBWorker.WorkerReportsProgress = true;
                TransactionBWorker.RunWorkerAsync();
            }
            else
            {
                Tracelog.TransWriteLine("Another process is running!");
                this.Close();
            }
        }

        private void TransactionBWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if(File.Exists(Application.StartupPath + @"\TransServices.log"))
            {
                File.Delete(Application.StartupPath + @"\TransServices.log");
            }
            using (var writer = File.AppendText(TxtIsrunning)) { };
            //SdkConnection.GetCompany();
            CreateTextCustomer.CreateCustomer();
        }

        private void TransactionBWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = string.Format("{0} % of {1} Data", e.ProgressPercentage, e.UserState);

        }

        private void TransactionBWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (File.Exists(TxtIsrunning))
            {
                File.Delete(TxtIsrunning);
            }
            this.Close();
        }
    }
}
