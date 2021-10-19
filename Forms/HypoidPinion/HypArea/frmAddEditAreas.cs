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
	public partial class frmAddEditAreas : Form
	{
		#region Variables
		public AreaModel areaModel = new AreaModel();
		private int _type;
		#endregion

		#region Method
		public frmAddEditAreas(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM KHU VỰC";
				_type = cGlobalVariables.Add;
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA KHU VỰC";
				_type = cGlobalVariables.Edit;
			}
		}
		void LoadListUsers() {
			DataTable dt = TextUtils.Select(cGlobalFunction.GetWithSelectionQuery(new string[] { "ID", "UserCode", "FullName", "BirthOfDate" }, "Users", "UserCode"));
			cbManager.Properties.DataSource = dt;
			cbManager.Properties.DisplayMember = "UserCode";
			cbManager.Properties.ValueMember = "ID";
		}

		void LoadDataToForm() {
			txbAreaCode.Text = areaModel.AreaCode;
			cbManager.EditValue = areaModel.ManagerID;
			txbAreaName.Text = areaModel.AreaName;
		}

		bool ValidateForm() {
			if (txbAreaCode.Text == "") {
				MessageBox.Show("Vui lòng nhập mã khu vực!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cbManager.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (AreaBO.Instance.CheckExist("AreaCode", txbAreaCode.Text) && _type == cGlobalVariables.Add)
			{
				MessageBox.Show("Mã khu vực đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (_type == cGlobalVariables.Edit && txbAreaCode.Text.Trim() != areaModel.AreaCode) {
				if (AreaBO.Instance.CheckExist("AreaCode", txbAreaCode.Text))
				{
					MessageBox.Show("Mã khu vực đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			return true;
		}

		bool SaveData() {
			if (ValidateForm()) {
				try {
					if (_type == cGlobalVariables.Add) {
						areaModel.AreaCode = txbAreaCode.Text.Trim();
						areaModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						areaModel.ManagerCode = cbManager.Text.Trim();
						areaModel.AreaName = txbAreaName.Text.Trim();

						AreaBO.Instance.Insert(areaModel);
						return true;
					}
					if (_type == cGlobalVariables.Edit) {
						areaModel.AreaCode = txbAreaCode.Text.Trim();
						areaModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						areaModel.ManagerCode = cbManager.Text.Trim();
						areaModel.AreaName = txbAreaName.Text.Trim();

						TextUtils.ExcuteProcedure("spHypUpdateArea",
									new string[] { "@id" , "@areaCode" , "@managerID" , "@managerCode", "@areaName" },
									new object[] { areaModel.ID, areaModel.AreaCode, areaModel.ManagerID, areaModel.ManagerCode , areaModel.AreaName});
						return true;
					}
				}
				catch (Exception ex)
				{
					return false;
				}
			}
			return false;
		}
		#endregion

		#region Event
		#endregion

		private void frmAddEditAreas_Load(object sender, EventArgs e)
		{
			LoadDataToForm();
			LoadListUsers();
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				areaModel = new AreaModel();
				this.Text = "THÊM KHU VỰC";
				_type = cGlobalVariables.Add;
				LoadDataToForm();
			}
		}

		private void btnSaveClose_Click(object sender, EventArgs e)
		{
			if (SaveData()) {
				this.DialogResult = DialogResult.OK;
			}
		}

		private void frmAddEditAreas_FormClosing(object sender, FormClosingEventArgs e)
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
