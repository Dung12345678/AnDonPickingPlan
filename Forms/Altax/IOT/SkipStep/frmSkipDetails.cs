using DevExpress.CodeParser;
using BMS.Business;
using BMS.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.util;
using System.Windows.Forms;

namespace BMS
{

	public partial class frmSkipDetails : _Forms
	{

		public SkipCDModel _skipCDModel = new SkipCDModel();

		public frmSkipDetails()
		{
			InitializeComponent();
		}
		private void frmProductWorkingDetail_Load(object sender, EventArgs e)
		{
			loadData();
		}

		/// <summary>
		/// Load dữ liệu vào form
		/// </summary>
		void loadData()
		{
			txtProduct.Text = Lib.ToString(_skipCDModel.Product);
			rtxtCDSkip.Text = Lib.ToString(_skipCDModel.CDSkip);
		}

		/// <summary>
		/// Validate trước khi cất dữ liệu
		/// </summary>
		/// <returns></returns>
		private bool ValidateForm()
		{
			string sql1 = "";
			if (_skipCDModel.ID > 0)
			{
				sql1 = $"SELECT top 1 ID from [SumitomoHyp].[dbo].[SkipCD] WHERE Product = '{txtProduct.Text.Trim()}' and ID <>'{_skipCDModel.ID}'";
			}
			else
			{
				sql1 = $"SELECT top 1 ID from [SumitomoHyp].[dbo].[SkipCD] WHERE Product = '{txtProduct.Text.Trim()}'";
			}
			int id = TextUtils.ToInt(TextUtils.ExcuteScalar(sql1));
			if (id > 0)
			{
				MessageBox.Show("Trùng nhóm sản phẩm", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}

		/// <summary>
		/// Cất dữ liệu
		/// </summary>
		/// <returns></returns>
		bool saveData()
		{
			if (!ValidateForm())
				return false;

			_skipCDModel.Product = Lib.ToString(txtProduct.Text.Trim());
			_skipCDModel.CDSkip = Lib.ToString(rtxtCDSkip.Text.Trim());

			if (_skipCDModel.ID > 0)
			{
				SkipCDBO.Instance.Update(_skipCDModel);
				return true;
			}
			else
			{
				SkipCDBO.Instance.Insert(_skipCDModel);
			}
			return true;
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
				_skipCDModel = new SkipCDModel();
				loadData();
			}
		}
		private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSave_Click(null, null);
		}

		private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void closeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}
