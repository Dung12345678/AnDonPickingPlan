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
	public partial class frmConfigShipTo : _Forms
	{
		public frmConfigShipTo()
		{
			InitializeComponent();
		}

		private void frmConfigShipToH_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			DataTable dt = TextUtils.Select("select * from ConfigShipTo");
			grdData.DataSource = dt;
		}
		private void btnNew_Click(object sender, EventArgs e)
		{
			frmConfigShipToDetail frm = new frmConfigShipToDetail();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = Lib.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ConfigShipToModel model = (ConfigShipToModel)ConfigShipToBO.Instance.FindByPK(id);
			frmConfigShipToDetail frm = new frmConfigShipToDetail();
			frm._ShipToModel = model;
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
					ConfigShipToBO.Instance.Delete(ID);
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
