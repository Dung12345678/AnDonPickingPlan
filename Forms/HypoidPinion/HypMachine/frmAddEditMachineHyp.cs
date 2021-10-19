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
	public partial class frmAddEditMachineHyp : _Forms
	{
		#region Variables
		public MachineModel machineModel = new MachineModel();
		private int _type;
		#endregion

		#region Method
		public frmAddEditMachineHyp(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM MÁY";
				_type = cGlobalVariables.Add;
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA MÁY";
				_type = cGlobalVariables.Edit;
			}
		}

		void LoadListStage()
		{
			DataTable dt = TextUtils.Select(cGlobalFunction.GetAllQuery("Stage", "StageCode"));
			cbStage.Properties.DataSource = dt;
			cbStage.Properties.DisplayMember = "StageCode";
			cbStage.Properties.ValueMember = "ID";
		}

		void LoadDataToForm()
		{
			txbMachineCode.Text = machineModel.MachineCode;
			cbStage.EditValue = machineModel.StageID;
			txbMachineName.Text = machineModel.MachineName;
			txbDescription.Text = machineModel.Description;
		}

		bool ValidateForm()
		{
			if (txbMachineCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã máy!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cbStage.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn công đoạn!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}


			if (MachineBO.Instance.CheckExist("MachineCode", txbMachineCode.Text))
			{
				if (_type == cGlobalVariables.Edit && txbMachineCode.Text.Trim() != machineModel.MachineCode || _type == cGlobalVariables.Add)
				{
					MessageBox.Show("Mã máy đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
						machineModel.MachineCode = txbMachineCode.Text.Trim();
						machineModel.StageID = TextUtils.ToInt(cbStage.EditValue);
						machineModel.StageCode = cbStage.Text.Trim();
						machineModel.MachineName = txbMachineName.Text.Trim();
						machineModel.Description = txbDescription.Text;

						MachineBO.Instance.Insert(machineModel);
						return true;
					}
					if (_type == cGlobalVariables.Edit)
					{
						machineModel.MachineCode = txbMachineCode.Text.Trim();
						machineModel.StageID = TextUtils.ToInt(cbStage.EditValue);
						machineModel.StageCode = cbStage.Text.Trim();
						machineModel.MachineName = txbMachineName.Text.Trim();
						machineModel.Description = txbDescription.Text;

						MachineBO.Instance.Update(machineModel);
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
		private void frmAddEditMachineHyp_Load(object sender, EventArgs e)
		{
			LoadListStage();
			LoadDataToForm();
		}

		private void frmAddEditMachineHyp_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				machineModel = new MachineModel();
				this.Text = "THÊM MÁY";
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

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveClose_Click(null, null);
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}
		#endregion

	}
}
