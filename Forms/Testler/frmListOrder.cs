using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmListOrder : _Forms
	{
		public frmListOrder()
		{
			InitializeComponent();
		}

		string _pathTester = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Tester");

		private void frmListOrder_Load(object sender, EventArgs e)
		{
			if (!Directory.Exists(_pathTester)) Directory.CreateDirectory(_pathTester);
			string[] arrFile = Directory.GetFiles(_pathTester);
			if(arrFile.Length > 0)
			{
				for (int i = 0; i < arrFile.Length; i++)
				{
					File.Delete(arrFile[i]);
				}
			}
				
			string statrtDate = dtpFromDate.Value.ToString("yyyy-MM-dd 00:00");
			string endDate = dtpEndDate.Value.ToString("yyyy-MM-dd 23:59");

			DataTable dt = TextUtils.GetDataTableFromSP("spGetOrderTestlerByDate", new string[] { "@StartDate", "@EndDate" }, new object[] { statrtDate, endDate });
			grdData.DataSource = dt;

		}

		private void dtpFromDate_ValueChanged(object sender, EventArgs e)
		{
			string statrtDate = dtpFromDate.Value.ToString("yyyy-MM-dd 00:00");
			string endDate = dtpEndDate.Value.ToString("yyyy-MM-dd 23:59");

			DataTable dt = TextUtils.GetDataTableFromSP("spGetOrderTestlerByDate", new string[] { "@StartDate", "@EndDate" }, new object[] {statrtDate, endDate});
			grdData.DataSource = dt;
		}

		private void dtpEndDate_ValueChanged(object sender, EventArgs e)
		{
			string statrtDate = dtpFromDate.Value.ToString("yyyy-MM-dd 00:00");
			string endDate = dtpEndDate.Value.ToString("yyyy-MM-dd 23:59");
			DataTable dt = TextUtils.GetDataTableFromSP("spGetOrderTestlerByDate", new string[] { "@StartDate", "@EndDate" }, new object[] { statrtDate, endDate });
			grdData.DataSource = dt;
		}

		private void grdData_Click(object sender, EventArgs e)
		{

		}

		private void grvData_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
		{

		}

		private void grvData_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
		{
			frmTestta2 frm = new frmTestta2();
			frm._Order = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
			frm.Show();
		}		
	}
}
