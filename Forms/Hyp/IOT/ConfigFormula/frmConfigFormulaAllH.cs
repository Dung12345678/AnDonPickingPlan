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
	public partial class frmConfigFormulaAllH : _Forms
	{
		public frmConfigFormulaAllH()
		{
			InitializeComponent();
		}

		private void frmConfigFormulaAllH_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			DataTable dt = LibIE.Select("select * from ConfigFormulaAll Order By ProductGroupCode, ProductTypeCode ,STT");
			grdData.DataSource = dt;
		}
		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = Lib.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ConfigFormulaAllModel model = (ConfigFormulaAllModel)ConfigFormulaAllBO.Instance.FindByPK(id);
			frmConfigFormulaAllDetailH frm = new frmConfigFormulaAllDetailH();
			frm.formulaModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			frmConfigFormulaAllDetailH frm = new frmConfigFormulaAllDetailH();
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

			string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProductTypeCode));

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa loại sản phẩm [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					ConfigFormulaAllBO.Instance.Delete(ID);
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
	}
}
