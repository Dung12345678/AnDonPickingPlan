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
	public partial class frmGoods : _Forms
	{

		public frmGoods()
		{
			InitializeComponent();
		}



		#region Events
		private void btnCreateKnife_Click(object sender, EventArgs e)
		{
			frmEditGoods frm = new frmEditGoods(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadGoods();
			}
		}

		private void btnEditKnife_Click(object sender, EventArgs e)
		{
			int ID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			frmEditGoods frm = new frmEditGoods(cGlobalVariables.Edit);
			GoodsModel goodsModel = (GoodsModel)GoodsBO.Instance.FindByPK(ID);
			frm._goodsModel = goodsModel;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadGoods();
			}
		}

		private void btnDelKnife_Click(object sender, EventArgs e)
		{
			int strID = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue(colID));
			string str = TextUtilsHP.ToString(grvData.GetFocusedRowCellValue(colCode));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa dụng cụ [{0}] không?", str), TextUtilsHP.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{

					GoodsBO.Instance.Delete(strID);
					grvData.DeleteSelectedRows();

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

		}
		void LoadGoods()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			DataTable dt = TextUtilsHP.GetDataTableFromSP("spLoadGoods", new string[] { "@Text", "@DateStart", "@DateEnd" }, new object[] { txtFindDate.Text.Trim(), dtpFrom.Value, dtpTo.Value });
			grdData.DataSource = dt;
		}

		#endregion

		private void btnFindDate_Click(object sender, EventArgs e)
		{
			LoadGoods();
		}

		private void frmGoods_Load(object sender, EventArgs e)
		{
			LoadGoods();
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportGoodsExcelH frm = new frmImportGoodsExcelH();
			frm.Show();
		}

		private void txtFindDate_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			LoadGoods();
		}
	}
}
