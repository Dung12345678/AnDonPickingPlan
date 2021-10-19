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
	public partial class frmConfigFormulaCD5H : _Forms
	{
		public frmConfigFormulaCD5H()
		{
			InitializeComponent();
		}
		private void frmConfigFormulaCD5_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			DataTable dt = LibIE.Select("select * from ConfigFormula");
			grdData.DataSource = dt;
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = Lib.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ConfigFormulaModel model = (ConfigFormulaModel)ConfigFormulaBO.Instance.FindByPK(id);
			ConfigFormulaDetailH frm = new ConfigFormulaDetailH();
			frm.formulaModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			ConfigFormulaDetailH frm = new ConfigFormulaDetailH();
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
					ConfigFormulaBO.Instance.Delete(ID);
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

		private void button1_Click(object sender, EventArgs e)
		{
			string ProductGroup = Lib.ToString(grvData.GetFocusedRowCellValue(colProductGroupCode));
			DataTable dt = LibIE.Select($"SELECT DISTINCT ProductTypeCode FROM dbo.ConfigFormula WHERE ProductGroupCode='{ProductGroup}'");
			string ProductTypeCode = "";
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				ProductTypeCode = Lib.ToString(dt.Rows[i]["ProductTypeCode"]);
				LibIE.ExcuteProcedure("spGetInsertFormula", new string[] { "@ProductTypeCode", "@ProductGroupCode" }, new object[] { ProductTypeCode, ProductGroup });

			}
		}
	}
}
