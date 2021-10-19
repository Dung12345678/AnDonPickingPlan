using System;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using DevExpress.Utils;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using DevExpress.XtraGrid.Columns;

namespace BMS
{
    public partial class frmThongKePhanTichDataH : _Forms
    {
        public frmThongKePhanTichDataH()
        {
            InitializeComponent();
        }

        private void frmThongKePhanTichData_Load(object sender, EventArgs e)
        {
            //dtpFromDate.Value = DateTime.Now.AddDays(-30);
            dtpFromDate.Value = DateTime.Now;
            dtpEndDate.Value = DateTime.Now;
        }

        void removeColumn()
        {
            for (int i = 0; i < grvData.Columns.Count; i++)
            {
                if (i < 1) continue;
                grvData.Columns[i].Visible = false;
                grvData.Columns[i].Caption = i.ToString();
                grvData.Columns[i].FieldName = "";
            }
        }

        void loadData()
        {
            removeColumn();
            DateTime from = new DateTime(dtpFromDate.Value.Year, dtpFromDate.Value.Month, dtpFromDate.Value.Day, 0, 0, 0);
            DateTime to = new DateTime(dtpEndDate.Value.Year, dtpEndDate.Value.Month, dtpEndDate.Value.Day, 23, 59, 59);

            if (txtFromOrder.Text.Trim() != "" && txtToOrder.Text.Trim() != "")
            {
                string sql = @"SELECT TOP 1 DateLR 
                    FROM dbo.ProductCheckHistoryDetail
                    WHERE OrderCode LIKE '%{0}%' ORDER BY DateLR";
                DateTime? date1 = TextUtils.ToDate2(LibIE.ExcuteScalar(string.Format(sql, txtFromOrder.Text.Trim())));
                DateTime? date2 = TextUtils.ToDate2(LibIE.ExcuteScalar(string.Format(sql, txtToOrder.Text.Trim())));
                if (date1.HasValue && date2.HasValue)
                {
                    if (date2.Value >= date1.Value)
                    {
                        from = new DateTime(date1.Value.Year, date1.Value.Month, date1.Value.Day, 0, 0, 0);
                        to = new DateTime(date2.Value.Year, date2.Value.Month, date2.Value.Day, 0, 0, 0);
                    }
                    else
                    {
                        to = new DateTime(date1.Value.Year, date1.Value.Month, date1.Value.Day, 0, 0, 0);
                        from = new DateTime(date2.Value.Year, date2.Value.Month, date2.Value.Day, 0, 0, 0);
                    }
                }
            }
          
            List<string> listColumnText = new List<string>();
            if (txtFindText.Text.Trim() == "")
            {
                return;
            }
            string[] arr = txtFindText.Text.Trim().Split(';');
            for (int i = 0; i < arr.Length; i++)
            {
                listColumnText.Add(string.Format("[{0}]", arr[i]));
                grvData.Columns[i + 1].VisibleIndex = i + 1;
                grvData.Columns[i + 1].Caption = arr[i];
                grvData.Columns[i + 1].FieldName = arr[i];
                grvData.Columns[i + 1].Width = 200;
            }

            DataTable dt = LibIE.LoadDataFromSP("spThongKePhanTichData", "A"
                , new string[] { "@From", "@To", "@FistColumnText" }
                , new object[] { from, to, string.Join(",", listColumnText) });

            //DataColumn c = new DataColumn();
            //c.Caption = "STT";
            //c.ColumnName = "STT";
            //c.DataType = typeof(int);
            //c.AutoIncrement = true;
            //c.AutoIncrementSeed = 1;
            //c.AutoIncrementStep = 1;
            //c.AllowDBNull = false;
            ////c.Unique = true;
            //dt.Columns.Add(c);
            //dt.Columns["STT"].SetOrdinal(0);
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
            if (grvData.RowCount < 0)
            {
                return;
            }

            TextUtils.ExportExcel(grvData);
        }
    }    
}
