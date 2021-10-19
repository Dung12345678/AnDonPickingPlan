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
	public partial class frmAddEditStageWorker : _Forms
	{
		#region Variables
		public StageWorkerModel stageWorkerModel = new StageWorkerModel();
		private int _type;
		#endregion


		#region Methods
		public frmAddEditStageWorker(int type)
		{
			InitializeComponent();
			if (type == cGlobalVariables.Add)
			{
				this.Text = "THÊM NHÂN VIÊN";
				_type = cGlobalVariables.Add;
			}
			if (type == cGlobalVariables.Edit)
			{
				this.Text = "SỬA THÔNG TIN";
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
			txbStageWorkerCode.Text = stageWorkerModel.WorkerCode;
			cbStage.EditValue = stageWorkerModel.StageID;
			txbJobDescription.Text = stageWorkerModel.JobDescription;
		}

		bool ValidateForm()
		{
			if (txbStageWorkerCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã công nhân!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cbStage.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn công đoạn!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}


			if (StageWorkerBO.Instance.CheckExist("WorkerCode", txbStageWorkerCode.Text))
			{
				if (_type == cGlobalVariables.Edit && txbStageWorkerCode.Text.Trim() != stageWorkerModel.WorkerCode || _type == cGlobalVariables.Add)
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
						stageWorkerModel.WorkerCode = txbStageWorkerCode.Text.Trim();
						stageWorkerModel.StageID = TextUtils.ToInt(cbStage.EditValue);
						stageWorkerModel.StageCode = cbStage.Text.Trim();
						stageWorkerModel.JobDescription = txbJobDescription.Text;
						StageBO.Instance.Insert(stageWorkerModel);
						return true;
					}
					if (_type == cGlobalVariables.Edit)
					{
						stageWorkerModel.WorkerCode = txbStageWorkerCode.Text.Trim();
						stageWorkerModel.StageID = TextUtils.ToInt(cbStage.EditValue);
						stageWorkerModel.StageCode = cbStage.Text.Trim();
						stageWorkerModel.JobDescription = txbJobDescription.Text;

						StageWorkerBO.Instance.Update(stageWorkerModel);
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

		#region Events
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
				stageWorkerModel = new StageWorkerModel();
				this.Text = "THÊM NHÂN VIÊN";
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

		
		private void frmAddEditStageWorker_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void frmAddEditStageWorker_Load(object sender, EventArgs e)
		{
			LoadListStage();
			LoadDataToForm();
		}
		#endregion

	}
}
