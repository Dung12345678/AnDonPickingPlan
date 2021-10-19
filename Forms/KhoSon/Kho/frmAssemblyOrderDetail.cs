using ST.Business;
using ST.Model;
using ST.Facade;
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
	public partial class frmAssemblyOrderDetail : _Forms
	{
		public OrderPartModel OrderPartModel = new OrderPartModel();
		private int type;

		public int Type
		{
			get
			{
				return type;
			}

			set
			{
				type = value;
			}
		}

		public frmAssemblyOrderDetail(int typeEvent)
		{
			InitializeComponent();
			Type = typeEvent;
			switch (typeEvent)
			{
				case 1:
					this.Text = "Thêm Order";
					break;
				case 2:
					this.Text = "Sửa Order";
					break;
			}
		}
		private bool ValidateForm()
		{
			if (txtOrderCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập Order Code.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtArticleID.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập mã Article ID.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;

		}
		void LoadData()
		{
			txtOrderCode.Text = OrderPartModel.OrderCode;
			txtArticleID.Text = OrderPartModel.ArticleID;
			txtCNT.Text = TextUtils.ToString(OrderPartModel.Cnt);
			txtDescription.Text = TextUtils.ToString(OrderPartModel.Description);
			txtQty.Text = TextUtils.ToString(OrderPartModel.Qty);
			txtShelf.Text = OrderPartModel.Shelf;
			txtLocation.Text = OrderPartModel.Location;
			txtLot.Text = OrderPartModel.Lot;
			txtUser.Text = OrderPartModel.Userr;
		}
		bool SaveData()
		{
			if (!ValidateForm()) return false;
			OrderPartModel.OrderCode = txtOrderCode.Text.Trim();
			OrderPartModel.ArticleID = txtArticleID.Text.Trim();
			OrderPartModel.Cnt = TextUtils.ToInt(txtCNT.Text.Trim());
			OrderPartModel.Description = txtDescription.Text.Trim();
			OrderPartModel.Qty = TextUtils.ToInt(txtQty.Text);
			OrderPartModel.Shelf = txtShelf.Text.Trim();
			OrderPartModel.Location = txtLocation.Text.Trim();
			OrderPartModel.Lot = txtLot.Text.Trim();
			OrderPartModel.Userr = txtUser.Text.Trim();
			OrderPartModel.OrderCodeAndCnt = OrderPartModel.OrderCode + OrderPartModel.Cnt;

			if (OrderPartModel.ID > 0)
			{
				OrderPartModel.UpdateDate = DateTime.Now;
				OrderPartModel.CreateAt = DateTime.Now;
				OrderPartBO.Instance.Update(OrderPartModel);
				MessageBox.Show("Update thành công").ToString();
			}
			else
			{
				OrderPartModel.CreateDate = DateTime.Now;
				OrderPartModel.CreateAt = DateTime.Now;
				OrderPartBO.Instance.Insert(OrderPartModel);
				MessageBox.Show("Thêm mới thành công").ToString();
			}
			return true;
		}

		private void frmAddEditPlan_Load(object sender, EventArgs e)
		{
			LoadData();
		}

		private void btnSaveClose_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				PartSonModel partSonModel = new PartSonModel();
				LoadData();
				this.Text = "Thêm kế hoạch";
				Type = 1;
			}
		}

		private void frmAddEditPlan_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveClose_Click(null, null);
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}
	}
}
