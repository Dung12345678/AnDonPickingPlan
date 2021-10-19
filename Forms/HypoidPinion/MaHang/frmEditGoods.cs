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
	public partial class frmEditGoods : _Forms
	{
		#region Variables
		public GoodsModel _goodsModel = new GoodsModel();
		#endregion
		#region Methods
		public frmEditGoods(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM MÃ HÀNG";
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA MÃ HÀNG";
			}
		}
		void LoadDataToForm()
		{
			txtGoodsName.Text = _goodsModel.Name;
			txtGoodsCode.Text = _goodsModel.Code;
		}

		bool ValidateForm()
		{
			if (txtGoodsCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã hàng!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtGoodsName.Text == "")
			{
				MessageBox.Show("Vui lòng nhập tên hàng!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (GoodsBO.Instance.CheckExist("Code", txtGoodsCode.Text.Trim()))
			{
				MessageBox.Show("Mã hàng đã tồn tại!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}

		bool SaveData()
		{
			if (ValidateForm())
			{
				try
				{
					_goodsModel.Code = txtGoodsCode.Text.Trim();
					_goodsModel.Name = txtGoodsName.Text.Trim();
					if (_goodsModel.ID > 0)
					{
						GoodsBO.Instance.Update(_goodsModel);
					}
					else
					{
						GoodsBO.Instance.Insert(_goodsModel);
					}
					return true;
				}
				catch (Exception ex)
				{
					return false;
				}
			}
			return false;
		}
		#endregion

		#region Events
		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveClose_Click(null, null);
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				_goodsModel = new GoodsModel();
				this.Text = "THÊM CÔNG CỤ";
				LoadDataToForm();
			}
		}

		private void btnSaveClose_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void frmAddEditKnife_Load(object sender, EventArgs e)
		{
			LoadDataToForm();
		}

		private void frmAddEditKnife_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
		#endregion
	}
}
