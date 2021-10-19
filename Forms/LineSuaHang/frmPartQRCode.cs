
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
	public partial class frmPartQRCode : _Forms
	{
		public frmPartQRCode()
		{
			InitializeComponent();
		}

		private void frmProductPlan_Load(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}
		void LoadInfoSearch()
		{
			DataTable dt = new DataTable();
			dt = TextUtilsStock.Select($"SELECT * FROM [ShiStock].[dbo].[Part]");
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
			frmPartQRCodeDetailH frm = new frmPartQRCodeDetailH();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			PartModel model = (PartModel)PartBO.Instance.FindByPK(id);
			frmPartQRCodeDetailH frm = new frmPartQRCodeDetailH();
			frm.partModel = model;
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

			string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("Name"));

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					PartBO.Instance.Delete(Convert.ToInt32(strID));
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
