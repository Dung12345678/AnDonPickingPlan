using System;
using System.Data;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmFindDataCD1H : _Forms
    {
        public frmFindDataCD1H()
        {
            InitializeComponent();
        }

        private void frmHistoryCheck_Load(object sender, EventArgs e)
        {
        }

        void loadData()
        {
            DataTable dt = LibIE.LoadDataFromSP("spFindDataCD1", "A"
                , new string[] { "@TextFind" }
                , new object[] { txtFindText.Text.Trim() });
            grdData.DataSource = dt;
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            loadData();
        }

        private void txtFindText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                loadData();
            }
        }

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0)
            {
                frmHistoryCheckDetailH frm = new frmHistoryCheckDetailH();
                frm.QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
                frm.Show();
            }
        }

        private void btnExportExecl_Click(object sender, EventArgs e)
        {
            if (grvData.RowCount <= 0)
            {
                return;
            }
            TextUtils.ExportExcel(grvData);
        }
    }    
}
