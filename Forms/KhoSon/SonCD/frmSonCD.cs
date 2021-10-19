using DevExpress.Utils;
using ST.Business;
using ST.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmSonCD : _Forms
	{

		public frmSonCD()
		{
			InitializeComponent();
		}

		private void frmProduct_Load(object sender, EventArgs e)
		{
			LoadSonCD();
		}
		void LoadSonCD()
		{
			DataTable dt = TextUtilsStock.GetDataTableFromSP("spGetSonCD", new string[] { "@TextFilter" }, new object[] { txtSeach.Text.Trim() });
			grdData.DataSource = dt;
		}

		private void btnFindAll_Click(object sender, EventArgs e)
		{
			LoadSonCD();
		}

		private void btnAdd_Click(object sender, EventArgs e)
		{
			frmSonCDDetails frm = new frmSonCDDetails();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadSonCD();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount == 0) return;
			int id = TextUtilsStock.ToInt(grvData.GetFocusedRowCellValue(colID));
			SonStepModel model = (SonStepModel)SonStepBO.Instance.FindByPK(id);
			frmSonCDDetails frm = new frmSonCDDetails();
			frm._sonStepModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadSonCD();
			}
		}

		private void btnDel_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount == 0)
				return;
			int strID = TextUtilsStock.ToInt(grvData.GetFocusedRowCellValue(colID));
			string str = TextUtilsStock.ToString(grvData.GetFocusedRowCellValue(colCDCode));

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					SonStepBO.Instance.Delete(strID);
					LoadSonCD();
				}
				catch
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
				}
			}
		}
	}
}
