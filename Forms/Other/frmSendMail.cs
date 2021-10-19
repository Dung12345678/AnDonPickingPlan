using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace BMS
{
    public partial class frmSendMail : _Forms
    {
        DataTable dtFile = new DataTable();
        int i = 0;

        public string To = "";
        public string Subject = "";
        public string Content = "";
        public string CC = "";

        public frmSendMail()
        {
            InitializeComponent();
        }

        private void frmSendMail_Load(object sender, EventArgs e)
        {
            txtContent.Text = Content;
            txtSubject.Text = Subject;
            txtTo.Text = To;
            txtCC.Text = CC;

            dtFile.Columns.Add("ID");
            dtFile.Columns.Add("FileName");
            dtFile.Columns.Add("Path");
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if (txtTo.Text.Trim() == "")
            {
                MessageBox.Show("Bạn hãy nhập địa chỉ mail đến!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            
            //bool success = TextUtils.SetEmailSend(txtSubject.Text.Trim(), txtContent.Text.Trim(), txtTo.Text.Trim(), txtCC.Text.Trim());

            #region Mở outlook
            int count = Process.GetProcesses().Where(o => o.ProcessName.Contains("OUTLOOK")).Count();
            if (count == 0)
            {
                try
                {
                    Process.Start("outlook.exe");
                }
                catch (Exception)
                {
                }
            }
            #endregion Mở outlook

            List<string> at = new List<string>();

            if (dtFile.Rows.Count > 0)
            {
                foreach (DataRow row in dtFile.Rows)
                {
                    at.Add(row["Path"].ToString());
                }
            }
            bool success = TextUtils.SetEmailSendHasAttach(txtSubject.Text.Trim(), txtContent.Text.Trim(), txtTo.Text.Trim(), txtCC.Text.Trim(), at);

            if (success)
            {
                MessageBox.Show("Đã gửi mail thành công!", TextUtils.Caption);
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Gửi không thành công!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnAttach_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string item in ofd.FileNames)
                {
                    i++;
                    dtFile.Rows.Add(i.ToString(), Path.GetFileName(item), item);
                }
                grvData.AutoGenerateColumns = false;
                grvData.DataSource = null;
                grvData.DataSource = dtFile;
            }
        }

        private void xóaFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (grvData.SelectedRows.Count == 0) return;
            int id = TextUtils.ToInt(grvData.SelectedRows[0].Cells[colID.Name].Value);
            if (id == 0) return;
            try
            {
                dtFile = dtFile.Select("ID <> '" + id + "'").CopyToDataTable();
            }
            catch (Exception)
            {
                dtFile.Rows.Clear();
            }
            grvData.AutoGenerateColumns = false;
            grvData.DataSource = null;
            grvData.DataSource = dtFile;
        }

        private void grvData_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string filePath in files)
                {
                    i++;
                    dtFile.Rows.Add(i.ToString(), Path.GetFileName(filePath), filePath);
                }
                grvData.AutoGenerateColumns = false;
                grvData.DataSource = null;
                grvData.DataSource = dtFile;
            }
        }

        private void grvData_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }
    }
}
