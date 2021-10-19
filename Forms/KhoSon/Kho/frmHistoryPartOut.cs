using BMS.Business;
using BMS.Model;
using ST.Business;
using ST.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmHistoryPartOut : _Forms
	{
		public frmHistoryPartOut()
		{
			InitializeComponent();
			//    dtgvHistory.AllowRestoreSelectionAndFocusedRow = DevExpress.Utils.DefaultBoolean.False;
		}
		private void frmHistoryPartOut_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			DataTable dataTable = TextUtilsStock.LoadDataFromSP(
						"spLoadHistoryPartOut"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@Text" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txbSearchHistory.Text.Trim()
				   });
			grdData.DataSource = dataTable;
		}

		private void btnSearchHistory_Click(object sender, EventArgs e)
		{
			LoadData();
		}
		private void dtpFrom_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearchHistory_Click(null, null);
			}
		}
		private void dtpTo_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearchHistory_Click(null, null);
			}
		}

		private void gvHistory_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0)
			{
				frmHistoryPartOutDetail frm = new frmHistoryPartOutDetail();
				frm._PartOutID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
				frm._OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				frm._Pid = TextUtils.ToString(grvData.GetFocusedRowCellValue(colPidAssembleStock));
				frm._StockCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colStockCode));
				frm.Show();
			}
		}

		private void btnXuatEx_Click(object sender, EventArgs e)
		{
			try
			{
				if (grvData.RowCount > 0)
				{
					FolderBrowserDialog od = new FolderBrowserDialog();
					if (od.ShowDialog() == DialogResult.OK)
					{
						TextUtilsStock.ExportExcel(grvData, od.SelectedPath, string.Format("Danh sách xuất kho{0}", DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private void txbSearchHistory_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			LoadData();
		}
	}
}
