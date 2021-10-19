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
using BMS.Model;

namespace BMS
{
    public partial class frmSendEmailAttach : _Forms
    {
        public DataTable dtFile = new DataTable();
        int i = 0;

        public bool IsClosed = false;
        public string To = "";
        public string Subject = "";
        public string Content = "";
        public string CC = "";
        public string Description = "";
        //public bool IsKCS = false;
        //public bool IsTK = false;
        //public ModuleErrorModel ModuleError = null;

        public frmSendEmailAttach()
        {
            InitializeComponent();
        }

        private void frmSendEmailAttach_Load(object sender, EventArgs e)
        {
            txtContent.Text = Content;
            txtSubject.Text = Subject;
            txtTo.Text = To;
            txtCC.Text = CC;

            webBrowser1.DocumentText = Content;

            if (dtFile.Columns.Count <= 0)
            {
                dtFile.Columns.Add("ID");
                dtFile.Columns.Add("FileName");
                dtFile.Columns.Add("Path");
            }

            grvData.AutoGenerateColumns = false;
            grvData.DataSource = dtFile;
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if (txtTo.Text.Trim() == "")
            {
                MessageBox.Show("Bạn hãy nhập địa chỉ mail đến!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

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

            if (txtDetail.Text.Trim()!="")
            {
                txtContent.Text += "<br>"
                + "<table style=\"width:100%\" border=\"1\">"
                 + " <tr>"
                   + " <td style=\"width:20%\"><b>Mô tả chi tiết</b></td>"
                    + "<td>" + txtDetail.Text.Trim().Replace("\n", "<br>") + "</td> "
                 + "</tr>"
                + "</table>";                
            }

            bool success = TextUtils.SetEmailSendHasAttach(txtSubject.Text.Trim(), txtContent.Text.Trim(), txtTo.Text.Trim(), txtCC.Text.Trim(), at);
            if (success)
            {
                MessageBox.Show("Đã gửi mail thành công!", TextUtils.Caption);
                Description = txtDetail.Text.Trim();
                this.DialogResult = DialogResult.OK;
                if (IsClosed)
                {
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Gửi không thành công!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void xóaFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
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

        private void grvData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string filePath = TextUtils.ToString(grvData.SelectedRows[0].Cells[colPath.Name].Value);
            if (filePath != "")
            {
                Process.Start(filePath);
            }
        }
    }
}
