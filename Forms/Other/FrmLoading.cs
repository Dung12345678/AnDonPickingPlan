using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
namespace BMS
{
    public partial class FrmLoading : Form
    {
        #region Fields
        BackgroundWorker _backgroundWorker;
        DateTime TimeStart;
        public bool IsCancelled = false;
        private bool _closeLoadingAfterCancel = false;
        #endregion

        #region Properties
        public string TextMessage
        {
            get { return lblTextMessage.Text; }
            set { lblTextMessage.Text = value; }
        }

        public bool ShowProgressBar
        {
            get { return progressBar1.Visible; }
            set { progressBar1.Visible = value; }
        }

        public int ProgressValue
        {
            get { return progressBar1.Value; }
            set { progressBar1.Value = value; }
        }
        public bool CloseLoadingAfterCancel
        {
            get 
            { 
                return _closeLoadingAfterCancel; 
            }
            set 
            {
                _closeLoadingAfterCancel = value; 
            }
        }
        public bool DisplayCancelButton
        {
            get
            {
                return btnCancel.Visible;
            }
            set
            {
                btnCancel.Visible = value;
            }
        }
        public BackgroundWorker BackgroundWorker
        {
            get { return _backgroundWorker; }
            set { _backgroundWorker = value; }
        }
        #endregion

        #region Constructors
        public FrmLoading()
        {
            InitializeComponent();
            TimeStart = DateTime.Now;
        }
        #endregion

        #region Events
        private void timer1_Tick(object sender, EventArgs e)
        {
            // if a backgroundworker is set, hide it if it's finish
            if (BackgroundWorker != null)
                if (!BackgroundWorker.IsBusy) this.Close();

            TimeSpan ts = DateTime.Now - TimeStart;
            label2.Text = ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds + "   " + ts.Milliseconds;
        }

        private void chkAlwaysontop_CheckedChanged(object sender, EventArgs e)
        {
            TopMost = chkAlwaysontop.Checked;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsCancelled = true;
            if (BackgroundWorker != null && BackgroundWorker.WorkerSupportsCancellation)
                BackgroundWorker.CancelAsync();
            if (CloseLoadingAfterCancel) Close();
        }


        #endregion

        #region Public Methods
        public void UpdateMessage(string newMessage)
        {
            TextMessage = newMessage;
        }
        #endregion

        #region Static Methods
        public static FrmLoading ShowLoading()
        {
            return ShowLoading(string.Empty);
        }

        public static FrmLoading ShowLoading(string message)
        {
            return ShowLoading(message, true, null);
        }

        public static FrmLoading ShowLoading(string message, bool closeLoadingAfterCancel, BackgroundWorker backgroundWorker)
        {
            FrmLoading newLoading = new FrmLoading();
            newLoading.TextMessage = message;
            newLoading.CloseLoadingAfterCancel = closeLoadingAfterCancel;
            newLoading.BackgroundWorker = backgroundWorker;
            newLoading.Show();
            return newLoading;
        }
        public static FrmLoading ShowLoading(string message, bool displayCancel, bool closeLoadingAfterCancel, BackgroundWorker backgroundWorker)
        {
            FrmLoading newLoading = new FrmLoading();
            newLoading.TextMessage = message;
            newLoading.CloseLoadingAfterCancel = closeLoadingAfterCancel;
            newLoading.BackgroundWorker = backgroundWorker;
            newLoading.DisplayCancelButton = displayCancel;
            newLoading.Show();
            return newLoading;
        }
        #endregion

        #region Extension Methods

        #endregion

        #region Private Methods

        #endregion

    }
}