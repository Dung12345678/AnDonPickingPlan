using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Utils;
using BMS.Model;
using BMS.Business;
using System.Collections;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System.Reflection;
using System.Threading;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using System.Diagnostics;
using DevExpress.XtraGrid.Localization;
using iTextSharp;
using iTextSharp.text.pdf;
using DevExpress.Utils;
using IOEx;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMS
{
    public partial class frmFindDataCD1 : _Forms
    {
        public frmFindDataCD1()
        {
            InitializeComponent();
        }

        private void frmHistoryCheck_Load(object sender, EventArgs e)
        {
        }

        void loadData()
        {
            DataTable dt = TextUtils.LoadDataFromSP("spFindDataCD1", "A"
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
                frmHistoryCheckDetail frm = new frmHistoryCheckDetail();
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
