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
	public partial class frmShelfOrGoodsDetail : _Forms
	{
		public ChangeShelfModel changeShelfModel = new ChangeShelfModel();
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

		public frmShelfOrGoodsDetail(int typeEvent)
		{
			InitializeComponent();
			Type = typeEvent;
			switch (typeEvent)
			{
				case 1:
					this.Text = "THÊM VỊ TRÍ";
					break;
				case 2:
					this.Text = "SỬA VỊ TRÍ";
					break;
			}
		}
		private bool ValidateForm()
		{
			if (txtShelf.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập vị trí.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtArticleID.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập mã Article ID.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			int check = 0;
			if (changeShelfModel.ID > 0)
			{
				check = TextUtilsStock.ToInt(TextUtilsStock.ExcuteScalar($"SELECT * FROM [ShiStock].[dbo].[ChangeShelf] WHERE ArticleID =N'{txtArticleID.Text.Trim()}' AND ID <> N'{changeShelfModel.ID}' "));
			}
			else
			{
				check = TextUtilsStock.ToInt(TextUtilsStock.ExcuteScalar($"SELECT * FROM [ShiStock].[dbo].[ChangeShelf] WHERE ArticleID =N'{txtArticleID.Text.Trim()}'"));
			}
			if (check > 0)
			{
				return false;
			}
			return true;

		}
		void LoadData()
		{
			txtArticleID.Text = changeShelfModel.ArticleID;
			txtShelf.Text = changeShelfModel.Shelf;
		}
		bool SaveData()
		{
			if (!ValidateForm()) return false;
			changeShelfModel.ArticleID = txtArticleID.Text.Trim();
			changeShelfModel.Shelf = txtShelf.Text.Trim();
			if (changeShelfModel.ID > 0)
			{
				ChangeShelfBO.Instance.Update(changeShelfModel);
				MessageBox.Show("Update thành công").ToString();
			}
			else
			{
				ChangeShelfBO.Instance.Insert(changeShelfModel);
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
				changeShelfModel = new ChangeShelfModel();
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
