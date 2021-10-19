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
	public partial class frmSkip : _Forms
	{
		public frmSkip()
		{
			InitializeComponent();
		}
		private void frmweightJig_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			DataTable dt = TextUtils.Select("select * from SkipCD");
			grdData.DataSource = dt;
		}
		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			SkipCDModel model = (SkipCDModel)SkipCDBO.Instance.FindByPK(id);
			frmSkipDetails frm = new frmSkipDetails();
			frm._skipCDModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			frmSkipDetails frm = new frmSkipDetails();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnDelete_Click(object sender, EventArgs e)
		{
			if (!grvData.IsDataRow(grvData.FocusedRowHandle))
				return;

			int ID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));

			string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProduct));

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa loại sản phẩm [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					SkipCDBO.Instance.Delete(ID);
					grvData.DeleteRow(grvData.FocusedRowHandle);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.\n" + ex.ToString());
				}
			}
		}

		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			btnEdit_Click(null, null);
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportSkipCDExcel frm = new frmImportSkipCDExcel();
			frm.Show();
		}
	}
}
