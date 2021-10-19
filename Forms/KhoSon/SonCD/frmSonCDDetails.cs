using ST.Business;
using ST.Model;
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
	public partial class frmSonCDDetails : _Forms
	{
		public SonStepModel _sonStepModel = new SonStepModel();
		public frmSonCDDetails()
		{
			InitializeComponent();
		}
		private void frmStockStepDetail_Load(object sender, EventArgs e)
		{
			txtStockCode.Text = _sonStepModel.SonStepCode;
			txtStockName.Text = _sonStepModel.SonStepName;
			txtStockCode.Focus();

		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				_sonStepModel = new SonStepModel();
				loadStockCDDetail();
			}
		}
		void loadStockCDDetail()
		{
			txtStockName.Text = _sonStepModel.SonStepName;
			txtStockCode.Text = _sonStepModel.SonStepCode;
			txtStockCode.Focus();
		}
		bool saveData()
		{
			if (!ValidateForm()) return false;
			_sonStepModel.SonStepName = txtStockName.Text.Trim();
			_sonStepModel.SonStepCode = txtStockCode.Text.Trim();
			if (_sonStepModel.ID > 0)
			{
				AssemblyStockBO.Instance.Update(_sonStepModel);
			}
			else
			{
				_sonStepModel.ID = (int)AssemblyStockBO.Instance.Insert(_sonStepModel);
			}
			return true;
		}
		private bool ValidateForm()
		{
			if (txtStockName.Text.Trim() == "" || txtStockCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập đầy đủ thông tin", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			string sql1 = "";
			if (_sonStepModel.ID > 0)
			{
				sql1 = $"SELECT top 1 ID from SonStep WHERE SonStepCode = N'{txtStockCode.Text.Trim()}' and ID <>'{_sonStepModel.ID}'";
			}
			else
			{
				sql1 = $"SELECT top 1 ID from SonStep WHERE SonStepCode = N'{txtStockCode.Text.Trim()}'";
			}
			int id = TextUtilsStock.ToInt(TextUtilsStock.ExcuteScalar(sql1));
			if (id > 0)
			{
				MessageBox.Show("Đã tồn tại mã công đoạn", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}

		private void KhaiBaoToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSave_Click(null, null);
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void frmStock_FormClosed(object sender, FormClosedEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
			//this.Dispose();
		}

		private void txtStockCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtStockName.Focus();
		}
	}
}
