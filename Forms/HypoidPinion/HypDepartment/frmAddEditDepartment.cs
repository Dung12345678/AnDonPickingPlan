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
	public partial class frmAddEditDepartment : _Forms
	{
		#region Variables
		public DepartmentModel departmentModel = new DepartmentModel();
		private int _type;
		#endregion

		#region Method
		public frmAddEditDepartment(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM BỘ PHẬN";
				_type = cGlobalVariables.Add;
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA BỘ PHẬN";
				_type = cGlobalVariables.Edit;
			}
		}

		void LoadListUsers()
		{
			DataTable dt = TextUtils.Select(cGlobalFunction.GetWithSelectionQuery(new string[] { "ID", "UserCode", "FullName", "BirthOfDate" }, "Users", "UserCode"));
			cbManager.Properties.DataSource = dt;
			cbManager.Properties.DisplayMember = "UserCode";
			cbManager.Properties.ValueMember = "ID";
		}

		void LoadListAreas()
		{
			DataTable dt = TextUtils.Select(cGlobalFunction.GetAllQuery("Area", "AreaCode"));
			cArea.Properties.DataSource = dt;
			cArea.Properties.DisplayMember = "AreaCode";
			cArea.Properties.ValueMember = "ID";
		}

		void LoadDataToForm() {
			txbDepartmentCode.Text = departmentModel.DepartmentCode;
			cbManager.EditValue = departmentModel.ManagerID;
			cArea.EditValue = departmentModel.AreaID;
			txbDepartmentName.Text = departmentModel.DepartmentName;
			txbDescription.Text = departmentModel.Description;
		}

		bool ValidateForm() {
			if (txbDepartmentCode.Text == "") {
				MessageBox.Show("Vui lòng nhập mã bộ phận!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cbManager.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cArea.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn khu vực!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (DepartmentBO.Instance.CheckExist("DepartmentCode", txbDepartmentCode.Text) && _type == cGlobalVariables.Add)
			{
				MessageBox.Show("Mã bộ phận đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (_type == cGlobalVariables.Edit && txbDepartmentCode.Text.Trim() != departmentModel.DepartmentCode)
			{
				if (DepartmentBO.Instance.CheckExist("DepartmentCode", txbDepartmentCode.Text))
				{
					MessageBox.Show("Mã khu vực đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			return true;
		}

		bool SaveData() {
			if (ValidateForm())
			{
				try
				{
					if (_type == cGlobalVariables.Add)
					{
						departmentModel.DepartmentCode = txbDepartmentCode.Text.Trim();
						departmentModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						departmentModel.AreaID = TextUtils.ToInt(cArea.EditValue);
						departmentModel.ManagerCode = cbManager.Text.Trim();
						departmentModel.AreaCode = cArea.Text.Trim();
						departmentModel.DepartmentName = txbDepartmentName.Text.Trim();
						departmentModel.Description = txbDescription.Text;

						DepartmentBO.Instance.Insert(departmentModel);
						return true;
					}
					if (_type == cGlobalVariables.Edit)
					{
						departmentModel.DepartmentCode = txbDepartmentCode.Text.Trim();
						departmentModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						departmentModel.AreaID = TextUtils.ToInt(cArea.EditValue);
						departmentModel.ManagerCode = cbManager.Text.Trim();
						departmentModel.AreaCode = cArea.Text.Trim();
						departmentModel.DepartmentName = txbDepartmentName.Text.Trim();
						departmentModel.Description = txbDescription.Text;

						TextUtils.ExcuteProcedure("spHypUpdateDepartment",
									new string[] { "@id", "@departmentCode", "@areaID", "@areaCode", "@managerID", "@managerCode", "@departmentName", "@description" },
									new object[] { departmentModel.ID, departmentModel.DepartmentCode, departmentModel.AreaID, departmentModel.AreaCode, departmentModel.ManagerID, departmentModel.ManagerCode, departmentModel.DepartmentName, departmentModel.Description });
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
		private void frmAddEditDepartment_Load(object sender, EventArgs e)
		{
			LoadListUsers();
			LoadListAreas();
			LoadDataToForm();

		}
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
				departmentModel = new DepartmentModel();
				this.Text = "THÊM BỘ PHẬN";
				_type = cGlobalVariables.Add;
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
		private void frmAddEditDepartment_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
		#endregion

	}
}
