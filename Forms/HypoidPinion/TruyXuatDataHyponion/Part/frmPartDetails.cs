using HP.Business;
using HP.Model;
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
	public partial class frmPartDetails : _Forms
	{
		public PartModel PartModel = new PartModel();
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

		public frmPartDetails(int typeEvent)
		{
			InitializeComponent();
			Type = typeEvent;
			switch (typeEvent)
			{
				case 1:
					this.Text = "THÊM BỘ PHẬN";
					break;
				case 2:
					this.Text = "SỬA BỘ PHẬN";
					break;
			}
		}
		private bool ValidateForm()
		{
			if (txtPartCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập mã bộ phận.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtPartName.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập tên bộ phận.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;

		}
		void LoadData()
		{
			txtPartCode.Text = PartModel.PartCode;
			txtPartName.Text = PartModel.PartName;
		}
		bool SaveData()
		{
			if (!ValidateForm()) return false;
			PartModel.PartCode = txtPartCode.Text.Trim();
			PartModel.PartName = txtPartName.Text.Trim();
			PartModel.CreateDate = DateTime.Now;

			if (PartModel.ID > 0)
			{
				PartBO.Instance.Update(PartModel);
				MessageBox.Show("Update thành công").ToString();
			}
			else
			{
				PartBO.Instance.Insert(PartModel);
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
