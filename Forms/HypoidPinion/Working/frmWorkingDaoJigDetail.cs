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
	public partial class frmWorkingDaoJigDetail : _Forms
	{
		// Loại = 0 Dao ; = 1 Jig ; = 2 Khuôn (để hiển thị nhóm ParentID == 0)
		public int _Loai;
		public int _GroupID;
		public int _StepID;
		public int IDTree;
		public WorkingModel _WorkingModel = new WorkingModel();

		public frmWorkingDaoJigDetail()
		{
			InitializeComponent();
		}
		void loadGroup()
		{
			// Load dụng cụ có parentID = 0 trong bảng KnifeDetailList
			DataTable dt = TextUtilsHP.GetDataTableFromSP("spLoadGroup", new string[] { "@Type" }, new object[] { _Loai });
			cboGroup.Properties.DataSource = dt;
			cboGroup.Properties.DisplayMember = "KnifeGroupCode";
			cboGroup.Properties.ValueMember = "ID";
		}
		/// <summary>
		/// Load dữ liệu vào form
		/// </summary>
		void loadData()
		{
			if (_WorkingModel.ID == 0)
			{
				txtSortOrder.Value = 1;
			}
			cboGroup.EditValue = _WorkingModel.KnifeGroupID;
			cboStep.EditValue = _WorkingModel.WorkingStepID;
			txtMax.Value = TextUtilsHP.ToDecimal(_WorkingModel.MaxValue);
			txtMin.Value = TextUtilsHP.ToDecimal(_WorkingModel.MinValue);
			txtName.Text = _WorkingModel.WorkingName;
			txtPeriodValue.Text = _WorkingModel.PeriodValue;
			txtSortOrder.Value = _WorkingModel.SortOrder;
			txtUnit.Text = _WorkingModel.Unit;
			cboValueType.SelectedIndex = _WorkingModel.ValueType;
			cboCheckValueType.SelectedIndex = _WorkingModel.CheckValueType;
			txtFrequency.Text = _WorkingModel.Frequency;
		}

		private void frmWorkingDetail_Load(object sender, EventArgs e)
		{
			
			loadGroup();
			loadData();
			if (_WorkingModel.ID == 0)
			{
				cboGroup.EditValue = _GroupID;
				cboStep.EditValue = _StepID;
			}
			txtName.Focus();
			if(_Loai==1)
			{
				txtFrequency.Visible = true;
				labelControl11.Visible = true;
			}
			else
			{
				labelControl11.Visible = false;
				txtFrequency.Visible = false;
			}	
		}
		/// <summary>
		/// Validate trước khi cất dữ liệu
		/// </summary>
		/// <returns></returns>
		private bool ValidateForm()
		{
			if (cboGroup.EditValue == null)
			{
				MessageBox.Show("Xin hãy chọn dụng cụ.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (cboStep.EditValue == null)
			{
				MessageBox.Show("Xin hãy chọn công đoạn.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cboValueType.SelectedIndex < 0)
			{
				MessageBox.Show("Xin hãy chọn kiểu kiểm tra.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (txtName.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy điền Mô tả công đoạn.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if(cboCheckValueType.SelectedIndex<=0)
			{
				MessageBox.Show("Xin hãy chọn một loại.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}	
			string sql1 = "";
			if (_WorkingModel.ID > 0)
			{
				sql1 = $"SELECT ID from [HypoidPinion].[dbo].[Working] WHERE SortOrder = '{txtSortOrder.Value}' and ID <>'{_WorkingModel.ID}' and WorkingStepID={cboStep.EditValue} ";
			}
			else
			{
				sql1 = $"SELECT ID from [HypoidPinion].[dbo].[Working] WHERE SortOrder = '{txtSortOrder.Value}' and WorkingStepID={cboStep.EditValue}";
			}
			int id = TextUtilsHP.ToInt(LibIE.ExcuteScalar(sql1));
			if (id > 0)
			{
				MessageBox.Show("Trùng STT", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				// return false;
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
			_WorkingModel.KnifeGroupID = TextUtilsHP.ToInt(cboGroup.EditValue);
			_WorkingModel.WorkingStepID = TextUtilsHP.ToInt(cboStep.EditValue);
			_WorkingModel.WorkingName = txtName.Text.Trim();
			_WorkingModel.SortOrder = TextUtilsHP.ToInt(txtSortOrder.Value);
			_WorkingModel.CheckValueType = cboCheckValueType.SelectedIndex;
			_WorkingModel.Unit = txtUnit.Text.Trim();
			_WorkingModel.ValueType = cboValueType.SelectedIndex;
			_WorkingModel.ValueTypeName = cboValueType.SelectedIndex == 0 ? "Check mark" : "Giá trị\n数値";
			_WorkingModel.MaxValue = TextUtilsHP.ToDouble(txtMax.Value);
			_WorkingModel.MinValue = TextUtilsHP.ToDouble(txtMin.Value);
			_WorkingModel.Frequency = txtFrequency.Text.Trim();
			if (_WorkingModel.ValueType == 0)//checkmark
			{
				_WorkingModel.PeriodValue = txtPeriodValue.Text.Trim();
			}
			else
			{
				if (_WorkingModel.CheckValueType == 1)//giá trị dạng số
				{
					if (_WorkingModel.MinValue == _WorkingModel.MaxValue)
					{
						_WorkingModel.PeriodValue = _WorkingModel.MaxValue.ToString("n3");
					}
					else
					{
						_WorkingModel.PeriodValue = _WorkingModel.MinValue.ToString("n3") + "~" + _WorkingModel.MaxValue.ToString("n3");
					}
				}
				else // Giá trị dạng ký tự
				{
					_WorkingModel.PeriodValue = txtPeriodValue.Text.Trim();
				}
			}

			if (_WorkingModel.ID > 0)
			{
				WorkingBO.Instance.Update(_WorkingModel);

				//Update những trường con theo trường cha
				TextUtilsHP.ExcuteSQL(string.Format("EXEC spUpdateMultiProductWorking {0}, 1", _WorkingModel.ID));
			}
			else
			{
				_WorkingModel.ID = (int)WorkingBO.Instance.Insert(_WorkingModel);
				//Insert những trường con theo trường cha
				TextUtilsHP.ExcuteSQL(string.Format("EXEC CreateKnifeWorking_ByWorkingID {0}", _WorkingModel.ID));
				//TextUtilsHP.ExcuteSQL(string.Format("EXEC CreateInsertKnifeJigWorking {0}", _knifeJigWorkingModel.ID));
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
				_WorkingModel = new WorkingModel();
				loadData();
				txtName.Focus();
			}
		}
		private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void txtMin_ValueChanged(object sender, EventArgs e)
		{
			txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
		}

		private void txtMax_ValueChanged(object sender, EventArgs e)
		{
			txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
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

		private void cboGroup_EditValueChanged(object sender, EventArgs e)
		{
			//Load Step
			loadStep();
		}
		void loadStep()
		{
			int groupID = TextUtils.ToInt(cboGroup.EditValue);
			DataTable dt = new DataTable();
			if (groupID > 0)
			{
				dt = LibIE.Select(string.Format("SELECT * FROM [HypoidPinion].[dbo].[WorkingStep] WHERE KnifeGroupID = {0} ORDER BY SortOrder", groupID));
			}
			cboStep.Properties.DataSource = dt;
			cboStep.Properties.DisplayMember = "WorkingStepCode";
			cboStep.Properties.ValueMember = "ID";
		}
	}
}
