using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmDataAccess : _Forms
	{
		public frmDataAccess()
		{
			InitializeComponent();
		}
		private void frmDataAccess_Load(object sender, EventArgs e)
		{

		}
		private void btnKnifeCode_Click(object sender, EventArgs e)
		{
			//Gửi mã Order và mã dao
			frmKnifeInformation frm = new frmKnifeInformation();
			frm._KnifeCode = btnKnifeCode.Text;
			frm._Goods = txtGoods.Text.Trim();
			TextUtils.OpenChildForm(frm, this);
		}

		private void btnJigCode_Click(object sender, EventArgs e)
		{
			frmJigLineInformation frm = new frmJigLineInformation();
			frm._JigCode = btnJigCode.Text.Trim();
			frm._Order = txtOrder.Text.Trim();
			frm._Goods = txtGoods.Text.Trim();
			TextUtils.OpenChildForm(frm, this);
		}

		private void btnMoldResult_Click(object sender, EventArgs e)
		{
			frmHypMoldLineInformation frm = new frmHypMoldLineInformation();
			frm.Order = txtOrder.Text.Trim();
			frm.Goods = txtGoods.Text.Trim();
			frm.Code = btnMoldResult.Text.Trim();
			TextUtils.OpenChildForm(frm, this);
		}

		private void txtCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				string[] array = txtCode.Text.Trim().Split(',');
				if (array.Count() <= 0 || array == null) return;
				string Goods = array[0];
				txtGoods.Text = Goods.Trim();
				if (array.Count() < 3) return;
				string Code = "";
				if (array.Count() == 5)
				{
					Code = array[3] + array[4];
				}
				else
				{
					Code = array[3];
				}
				DataSet ds = TextUtilsHP.GetListDataFromSP("spGetDataAccess", new string[] { "@Goods", "@Code" }, new object[] { Goods, Code });

				DataTable dtOrder = ds.Tables[0];
				DataTable dtDao = ds.Tables[1];
				DataTable dtJig = ds.Tables[2];
				DataTable dtKhuon = ds.Tables[3];
				if (dtOrder.Rows.Count > 0)
				{
					txtOrder.Text = TextUtils.ToString(dtOrder.Rows[0]["OrderMachining"]);
				}
				if (dtDao.Rows.Count>0)
				{
					btnKnifeCode.Text = TextUtils.ToString(dtDao.Rows[0]["KnifeCode"]);
				}
				if (dtJig.Rows.Count > 0)
				{
					btnJigCode.Text = TextUtils.ToString(dtJig.Rows[0]["KnifeDetailCode"]);
				}
				if (dtKhuon.Rows.Count > 0)
				{
					btnMoldResult.Text = TextUtils.ToString(dtKhuon.Rows[0]["KnifeDetailCode"]);
				}

			}
		}
	}
}
