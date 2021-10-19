using BMS.Business;
using BMS.Model;
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
	public partial class frmProductPlan : _Forms
	{
		public frmProductPlan()
		{
			InitializeComponent();
		}

		private void frmProductPlan_Load(object sender, EventArgs e)
		{
			dtpFrom.Value = DateTime.Now;//.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = DateTime.Now;//.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			LoadInfoSearch();
		}
		void LoadInfoSearch()
		{
			DataTable dt = new DataTable();
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			dt = TextUtils.LoadDataFromSP(
					   "spGetProductionPlanFindDate"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@TextFilter", "@IsShow","@UnFinished" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txtAssyOrderId.Text.Trim()
										, chkIsShow.Checked
										, chkUnFinished.Checked}
				   );

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
			frmProductPlanDetail frm = new frmProductPlanDetail();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ProductionPlanModel model = (ProductionPlanModel)ProductionPlanBO.Instance.FindByPK(id);
			frmProductPlanDetail frm = new frmProductPlanDetail();
			frm.ProductionPlan = model;
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
					ProductionPlanBO.Instance.Delete(Convert.ToInt32(strID));
					LoadInfoSearch();
				}
				catch
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");

				}
			}
		}

		private void btnImportExcelDate_Click(object sender, EventArgs e)
		{
			frmImportAssemblyDateExcelAltax frm = new frmImportAssemblyDateExcelAltax();
			frm.Show();
		}

		private void btnRestart_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void grvData_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
			//Status = 1 Đã xong xanh
			try
			{
				string Statuss = Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colStatuss));
				if (Statuss.ToUpper().Contains("ĐÃ XONG"))
				{
					e.Appearance.BackColor = Color.Lime;
				}
				if (Statuss.ToUpper().Contains("QUÁ HẠN"))
				{
					e.Appearance.BackColor = Color.Yellow;
				}
			}
			catch
			{

			}
		}

		private void chkUnFinished_CheckedChanged(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void txtAssyOrderId_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			LoadInfoSearch();
		}

		private void btnExcel_Click(object sender, EventArgs e)
		{
			try
			{
				if (grvData.RowCount > 0)
				{
					FolderBrowserDialog od = new FolderBrowserDialog();
					if (od.ShowDialog() == DialogResult.OK)
					{
						TextUtils.ExportExcel(grvData, od.SelectedPath, string.Format("Kế hoạch sản xuất Altax_{0}", DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
	}
}
