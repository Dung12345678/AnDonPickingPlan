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
	public partial class frmAddEditStageHyp : _Forms
	{
		#region Variables
		public StageModel stageModel = new StageModel();
		private int _type;
		#endregion

		#region Method
		public frmAddEditStageHyp(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM CÔNG ĐOẠN";
				_type = cGlobalVariables.Add;
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA CÔNG ĐOẠN";
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

		void LoadListDepartment()
		{
			DataTable dt = TextUtils.Select(cGlobalFunction.GetAllQuery("Department", "DepartmentCode"));
			cDepartment.Properties.DataSource = dt;
			cDepartment.Properties.DisplayMember = "DepartmentCode";
			cDepartment.Properties.ValueMember = "ID";
		}

		void LoadDataToForm()
		{
			txbStageCode.Text = stageModel.StageCode;
			cbManager.EditValue = stageModel.ManagerID;
			cDepartment.EditValue = stageModel.DepartmentID;
			txbStageName.Text = stageModel.StageName;
			txbDescription.Text = stageModel.Description;
		}

		bool ValidateForm()
		{
			if (txbStageCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã công đoạn!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cbManager.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cDepartment.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn bộ phận!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			
			if (StageBO.Instance.CheckExist("StageCode", txbStageCode.Text))
			{
				if (_type == cGlobalVariables.Edit && txbStageCode.Text.Trim() != stageModel.StageCode || _type == cGlobalVariables.Add)
				{
					MessageBox.Show("Mã công đoạn đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			return true;
		}

		bool SaveData()
		{
			if (ValidateForm())
			{
				try
				{
					if (_type == cGlobalVariables.Add)
					{
						stageModel.StageCode= txbStageCode.Text.Trim();
						stageModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						stageModel.DepartmentID= TextUtils.ToInt(cDepartment.EditValue);
						stageModel.ManagerCode = cbManager.Text.Trim();
						stageModel.DepartmentCode = cDepartment.Text.Trim();
						stageModel.StageName = txbStageName.Text.Trim();
						stageModel.Description = txbDescription.Text;

						StageBO.Instance.Insert(stageModel);
						return true;
					}
					if (_type == cGlobalVariables.Edit)
					{
						stageModel.StageCode = txbStageCode.Text.Trim();
						stageModel.ManagerID = TextUtils.ToInt(cbManager.EditValue);
						stageModel.DepartmentID = TextUtils.ToInt(cDepartment.EditValue);
						stageModel.ManagerCode = cbManager.Text.Trim();
						stageModel.DepartmentCode = cDepartment.Text.Trim();
						stageModel.StageName = txbStageName.Text.Trim();
						stageModel.Description = txbDescription.Text;

						TextUtils.ExcuteProcedure("spHypUpdateStage",
									new string[] { "@id", "@stageCode", "@departmentID", "@departmentCode", "@managerID", "@managerCode", "@stageName", "@description" },
									new object[] { stageModel.ID, stageModel.StageCode, stageModel.DepartmentID, stageModel.DepartmentCode, stageModel.ManagerID, stageModel.ManagerCode, stageModel.StageName, stageModel.Description });
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
				stageModel = new StageModel();
				this.Text = "THÊM CÔNG ĐOẠN";
				_type = cGlobalVariables.Add;
				LoadDataToForm();
			}
		}

		private void frmAddEditStageHyp_Load(object sender, EventArgs e)
		{
			LoadListUsers();
			LoadListDepartment();
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

		private void frmAddEditStageHyp_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;

		}
		#endregion

	}
}
