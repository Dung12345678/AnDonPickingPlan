using BMS;
using DevExpress.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmChooseBackupVersion : _Forms
    {
        frmCshLoading _frmLoad = null;
        public frmChooseBackupVersion()
        {
            InitializeComponent();
        }

        private void frmChooseBackupVersion_Load(object sender, EventArgs e)
        {
            loadData();
        }

        void loadData()
        {
            DataTable dt = TextUtils.Select($"select ROW_NUMBER() OVER(ORDER BY ID desc) AS STT, * from SumitomoTest.dbo.BackupVersion");
            grdData.DataSource = dt;
        }
        private void btnChoose_Click(object sender, EventArgs e)
        {
            //Nếu đang sử dụng database thì thôi
            if (TextUtils.ToBoolean(grvData.GetFocusedRowCellValue(colIsUse)))
            {
                MessageBox.Show("Dữ liệu này đang được sử dụng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            //progressBar.Visible = true;
            //Thread.Sleep(5000);
            //progressBar.Visible = false;
            //return;
            using (WaitDialogForm fWait = new WaitDialogForm("", ""))
            {
                try
                {
                    string fileAltax = TextUtils.ToString(grvData.GetFocusedRowCellValue(colAltax));
                    string fileHyp = TextUtils.ToString(grvData.GetFocusedRowCellValue(colHyp));
                    int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));

                    //Restore database file đã chọn vào database backup
                    string sqlRestore = $"Use [SumitomoTest]; exec spRestoreDatabase @FileNameAltax = '{fileAltax.Trim()}',@FileNameHyp='{fileHyp.Trim()}'";
                    //Update lại trạng thái database đã được sử dụng
                    string sqlUpdateIsUse = $"Use [SumitomoTest]; exec spUpdateBackupVersion {id}";

                    TextUtils.ExcuteSQL($"{sqlRestore};{sqlUpdateIsUse}");

                    loadData();

                    this.DialogResult = DialogResult.OK;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}
