
using ST.Business;
using ST.Model;
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
	public partial class frmErrorDetail : _Forms
	{
		public ErrorModel errorModel = new ErrorModel();
		public int ProductGroupID { get; set; }

		public frmErrorDetail()
		{
			InitializeComponent();
		}
		private void frmWorkingDetail_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			txtCode.Text = errorModel.Code;
			txtName.Text = errorModel.Name;
		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			if (saveData())
				this.DialogResult = DialogResult.OK;
		}
		bool saveData()
		{
			errorModel.Code = txtCode.Text.Trim();
			errorModel.Name = txtName.Text.Trim();
			if (errorModel.ID > 0)
			{
				ErrorBO.Instance.Update(errorModel);
			}
			else
			{
				ErrorBO.Instance.Insert(errorModel);
			}
			return true;
		}
		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				errorModel = new ErrorModel();
				txtCode.Text = "";
				txtName.Text = "";
				txtCode.Focus();
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
