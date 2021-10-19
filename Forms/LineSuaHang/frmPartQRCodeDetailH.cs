
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
	public partial class frmPartQRCodeDetailH : _Forms
	{
		public PartModel partModel = new PartModel();
		public int ProductGroupID { get; set; }

		public frmPartQRCodeDetailH()
		{
			InitializeComponent();
		}
		private void frmWorkingDetail_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			txtCode.Text = partModel.Code;
			txtName.Text = partModel.Name;
		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			if (saveData())
				this.DialogResult = DialogResult.OK;
		}
		bool saveData()
		{
			partModel.Code = txtCode.Text.Trim();
			partModel.Name = txtName.Text.Trim();
			if (partModel.ID > 0)
			{
				PartBO.Instance.Update(partModel);
			}
			else
			{
				PartBO.Instance.Insert(partModel);
			}
			return true;
		}
		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				partModel = new PartModel();
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
