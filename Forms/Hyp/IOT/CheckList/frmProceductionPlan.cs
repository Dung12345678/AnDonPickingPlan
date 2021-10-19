using IE;
using IE.Business;
using IE.Model;
using IE.Utils;
using DevExpress.Data.Filtering.Helpers;
using System;
using System.Collections;
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
	public partial class frmProceductionPlan : _Forms
	{
		public frmProceductionPlan()
		{
			InitializeComponent();
		}
		// khởi tạo delegate để chuyển dữ liệu form này sang from detail
		//  public delegate void delPassData(TextBox text);
		private void btnNew_Click(object sender, EventArgs e)
		{
			frmProductionPlanDetail frm = new frmProductionPlanDetail();
			frm.Show();
		}
		private void setInterface(bool isEdit)
		{
			grdData.Enabled = !isEdit;
			btnEdit.Visible = !isEdit;
		}

		private void grdData_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0 && btnEdit.Enabled == true)
				btnEdit_Click(null, null);
		}
		int _rownIndex = 0;

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ProductionPlanModel model = (ProductionPlanModel)ProductionPlanBO.Instance.FindByPK(id);
			frmProductionPlanDetail frm = new frmProductionPlanDetail();
			frm.ProductionPlan = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}

		}
		private void frmProceductionPlan_Load(object sender, EventArgs e)
		{
			dtpFrom.Value = DateTime.Now.AddDays(-30);//.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = DateTime.Now;//.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			LoadInfoSearch();
		}


		private void btnAdd_Click(object sender, EventArgs e)
		{
			frmProductionPlanDetail frm = new frmProductionPlanDetail();
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

		private void grdData_Click(object sender, EventArgs e)
		{
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportProductionPlanExcel frm = new frmImportProductionPlanExcel();
			frm.Show();
		}
		void LoadInfoSearch()
		{
			DataTable dt = new DataTable();
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			dt = LibIE.LoadDataFromSP(
					   "spGetProductionPlanFindDate"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@TextFilter", "@IsShow", "@UnFinished" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txtAssyOrderId.Text.Trim()
										, chkIsShow.Checked
										, chkUnFinished.Checked}
				   );

			grdData.DataSource = dt;
		}

		private void btnFindDate_Click(object sender, EventArgs e)
		{

			LoadInfoSearch();
		}

		private void dtpFrom_ValueChanged(object sender, EventArgs e)
		{

		}
		private void btnRestart_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void btnImportExcelDate_Click(object sender, EventArgs e)
		{
			frmImportAssemblyDateExcel frm = new frmImportAssemblyDateExcel();
			frm.Show();

		}
		private void grvData_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
			//Ngày lắp ráp có thì đổi màu
			//string AssemblyDate = Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colAssemblyDate));
			//if (AssemblyDate != "")//Điều kiện
			//{
			//	e.Appearance.BackColor = Color.Orange;
			//}
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

		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			btnEdit_Click(null, null);
		}

		private void chkIsShow_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void chkUnFinished_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void txtAssyOrderId_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			btnFindDate_Click(null, null);
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
						TextUtils.ExportExcel(grvData, od.SelectedPath, string.Format("Kế hoạch sản xuất Hyp_{0}", DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
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
