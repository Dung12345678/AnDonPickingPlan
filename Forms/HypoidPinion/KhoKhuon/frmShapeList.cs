using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HP.Business;
using HP.Model;

namespace BMS
{
	public partial class frmShapeList : _Forms
	{
		#region Variables
		int prevRow;
		#endregion

		#region Methods
		public frmShapeList()
		{
			InitializeComponent();
		}
		private void frmJigList_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			DataTable arr = TextUtilsHP.GetDataTableFromSP(
											"spLoadShapeList",
											new string[] { "@DateStart", "@DateEnd", "@Text" },
											new object[]{dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
														, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
														, txtFindDate.Text.Trim()});
			grdData.DataSource = arr;
		}
		#endregion

		#region Events
		private void btnCreateJig_Click(object sender, EventArgs e)
		{
			frmShapeDetails frm = new frmShapeDetails();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnEditJig_Click(object sender, EventArgs e)
		{
			int id = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			prevRow = grvData.GetSelectedRows()[0];
			KnifeDetailListModel model = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(id);

			frmShapeDetails frm = new frmShapeDetails();
			frm._KnifeDetailListModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
				grvData.FocusedRowHandle = prevRow;

			}
		}

		private void btnDelJig_Click(object sender, EventArgs e)
		{
			if (!grvData.IsDataRow(grvData.FocusedRowHandle))
				return;
			int strID = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue(colID));
			string str = TextUtilsHP.ToString(grvData.GetFocusedRowCellValue(colKnifeCode));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa dao [{0}] không?", str), TextUtilsHP.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					//if (KnifeSharpeningDetailsBO.Instance.CheckExist("KnifeID", strID))
					//{
					//	MessageBox.Show("Dao đã được sử dụng! Vui lòng chọn hủy dao thay cho xóa", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//	return;
					//}
					//else
					//{
					KnifeDetailListBO.Instance.Delete(strID);
					grvData.DeleteRow(grvData.FocusedRowHandle);
					//}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void btnRefresh_Click(object sender, EventArgs e)
		{
			LoadData();
		}
		private void btnFindDate_Click(object sender, EventArgs e)
		{
			LoadData();
		}
		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			btnEditJig_Click(null, null);
		}
		#endregion


	}
}
