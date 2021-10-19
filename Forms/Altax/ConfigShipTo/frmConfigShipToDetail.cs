using BMS.Business;
using BMS.Model;
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
	public partial class frmConfigShipToDetail : _Forms
	{
		public ConfigShipToModel _ShipToModel = new ConfigShipToModel();
		public frmConfigShipToDetail()
		{
			InitializeComponent();
		}
		private void frmConfigShipToDetailH_Load(object sender, EventArgs e)
		{
			loadConfigShipToDetail();
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
				_ShipToModel = new ConfigShipToModel();
				loadConfigShipToDetail();
			}
		}
		void loadConfigShipToDetail()
		{
			txtShipTo.Text = _ShipToModel.ShipTo;
			chkIsHidden.Checked = _ShipToModel.IsHidden;
		}
		private bool ValidateForm()
		{
			if (txtShipTo.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập Đích", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			string sql1 = "";
			if (_ShipToModel.ID > 0)
			{
				sql1 = $"SELECT top 1 ID from ConfigShipTo WHERE ShipTo = N'{txtShipTo.Text.Trim()}' and ID <>'{_ShipToModel.ID}'";
			}
			else
			{
				sql1 = $"SELECT top 1 ID from ConfigShipTo WHERE ShipTo = N'{txtShipTo.Text.Trim()}'";
			}
			int id = TextUtils.ToInt(TextUtils.ExcuteScalar(sql1));
			if (id > 0)
			{
				MessageBox.Show("Đã tồn tại Đích", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}
		bool saveData()
		{
			if (!ValidateForm()) return false;
			_ShipToModel.ShipTo = txtShipTo.Text.Trim();
			_ShipToModel.IsHidden = chkIsHidden.Checked;
			if (_ShipToModel.ID > 0)
			{
				ConfigShipToBO.Instance.Update(_ShipToModel);
			}
			else
			{
				_ShipToModel.ID = (int)ConfigShipToBO.Instance.Insert(_ShipToModel);
			}
			return true;
		}
		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSave_Click(null, null);
		}

		private void closeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
		private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void frmConfigShipToDetailH_FormClosed(object sender, FormClosedEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}
