using IE.Business;
using IE.Model;
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
	public partial class frmReduceSpeedH : _Forms
	{
		public frmReduceSpeedH()
		{
			InitializeComponent();
		}

		private void frmProductPlan_Load(object sender, EventArgs e)
		{
			dtpFrom.Value = DateTime.Now.AddMonths(-1);
			dtpTo.Value = DateTime.Now;
			LoadInfoSearch();
		}
		void LoadInfoSearch()
		{
			DataTable dt = new DataTable();
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			dt = LibIE.LoadDataFromSP(
					   "spGetSpeedReducer"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@TextFilter" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txtName.Text.Trim()
										});

			grdData.DataSource = dt;
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportProductPlanExcel frm = new frmImportProductPlanExcel();
			frm.Show();
		}

		private void btnFindDate_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void btnAdd_Click(object sender, EventArgs e)
		{
			frmReduceSpeedDetailH frm = new frmReduceSpeedDetailH();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			SpeedReducerModel model = (SpeedReducerModel)SpeedReducerBO.Instance.FindByPK(id);
			frmReduceSpeedDetailH frm = new frmReduceSpeedDetailH();
			frm._SpeedReducerModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}

		private void btnDel_Click(object sender, EventArgs e)
		{
			if (!grvData.IsDataRow(grvData.FocusedRowHandle))
				return;

			int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));

			string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("ProductCode"));

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					SpeedReducerBO.Instance.Delete(Convert.ToInt32(strID));
					LoadInfoSearch();
				}
				catch
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");

				}
			}
		}
		private void btnRestart_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void grvData_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
	
		}
		private void txtAssyOrderId_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			LoadInfoSearch();
		}
	}
}
