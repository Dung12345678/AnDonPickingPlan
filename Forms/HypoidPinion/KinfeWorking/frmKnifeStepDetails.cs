using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmKnifeStepDetails : _Forms
	{
		public KnifeStepModel _KnifeStep = new KnifeStepModel();

		public int _WorkingStepID = 0;
		public int _KnifeDetailListID = 0;

		public frmKnifeStepDetails()
		{
			InitializeComponent();
		}
		/// <summary>
		/// Load dữ liệu vào form
		/// </summary>
		void loadData()
		{
			txtSortOrder.Value = _KnifeStep.SortOrder;
			txtCode.Text = _KnifeStep.ProductStepCode;
			txtDes.Text = _KnifeStep.Description;
		}

		private void frmProductStepDetail_Load(object sender, EventArgs e)
		{
			loadData();
		}

		/// <summary>
		/// Validate trước khi cất dữ liệu
		/// </summary>
		/// <returns></returns>
		private bool ValidateForm()
		{
			if (txtSortOrder.Value <= 0)
			{
				MessageBox.Show("Xin hãy điền Thứ tự của công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (txtCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy điền Mã cho công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (txtDes.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy điền Mô tả cho công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
		//	_KnifeStep.WorkingStepID = _WorkingStepID;
			_KnifeStep.KnifeDetailListID = _KnifeDetailListID;
			_KnifeStep.ProductStepCode = txtCode.Text.Trim();
			_KnifeStep.SortOrder = (int)txtSortOrder.Value;
			_KnifeStep.Description = txtDes.Text.Trim();
			_KnifeStep.UpdatedBy = Global.AppUserName;
			_KnifeStep.UpdatedDate = DateTime.Now;

			if (_KnifeStep.ID <= 0)
			{
				_KnifeStep.CreatedBy = Global.AppUserName;
				_KnifeStep.CreatedDate = DateTime.Now;

				KnifeStepBO.Instance.Insert(_KnifeStep);
			}
			else
			{
				KnifeStepBO.Instance.Update(_KnifeStep);
			}
			return true;
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			if (saveData())
				this.DialogResult = DialogResult.OK;
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				_KnifeStep = new KnifeStepModel();
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

		}

		private void closeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}
