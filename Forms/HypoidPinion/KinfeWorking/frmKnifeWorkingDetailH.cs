using DevExpress.CodeParser;
using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.util;
using System.Windows.Forms;

namespace BMS
{

	public partial class frmKnifeWorkingDetailH : _Forms
	{
		public KnifeJigWorkingModel _KnifeJigWorking = new KnifeJigWorkingModel();
		public int StepID = 0;
		public int KnifeDetailListID = 0;
		public int Loai;
		public frmKnifeWorkingDetailH()
		{
			InitializeComponent();
		}

		private void frmProductWorkingDetail_Load(object sender, EventArgs e)
		{
			loadStep();
			loadData();
			if (_KnifeJigWorking.ID == 0)
			{
				cboStep.EditValue = StepID;
			}
			txtName.Focus();
		}
		void loadStep()
		{
			if (Loai == 3)
			{
				DataTable dt = new DataTable();
				dt = TextUtilsHP.Select(string.Format("SELECT * FROM [HypoidPinion].[dbo].[KnifeStep] WHERE ToolsID = {0} ORDER BY SortOrder", KnifeDetailListID));
				cboStep.Properties.DataSource = dt;
				cboStep.Properties.DisplayMember = "ProductStepCode";
				cboStep.Properties.ValueMember = "ID";
			}
			else
			{
				DataTable dt = new DataTable();
				dt = TextUtilsHP.Select(string.Format("SELECT * FROM [HypoidPinion].[dbo].[KnifeStep] WHERE KnifeDetailListID = {0} ORDER BY SortOrder", KnifeDetailListID));
				cboStep.Properties.DataSource = dt;
				cboStep.Properties.DisplayMember = "ProductStepCode";
				cboStep.Properties.ValueMember = "ID";
			}
		}

		/// <summary>
		/// Load dữ liệu vào form
		/// </summary>
		void loadData()
		{
			if (_KnifeJigWorking.ID == 0)
			{
				txtSortOrder.Value = 1;
			}
			else
			{
				txtSortOrder.Value = _KnifeJigWorking.SortOrder;
			}
			txtMax.Value = _KnifeJigWorking.MaxValue;
			txtMin.Value = _KnifeJigWorking.MinValue;
			txtName.Text = _KnifeJigWorking.WorkingName;
			txtPeriodValue.Text = _KnifeJigWorking.PeriodValue;
			txtUnit.Text = _KnifeJigWorking.Unit;
			cboValueType.SelectedIndex = _KnifeJigWorking.ValueType;
			cboStep.EditValue = _KnifeJigWorking.KnifeStepID;
			cboCheckValueType.SelectedIndex = _KnifeJigWorking.CheckValueType;
			txtFrequency.Text = _KnifeJigWorking.Frequency;
		}

		/// <summary>
		/// Validate trước khi cất dữ liệu
		/// </summary>
		/// <returns></returns>
		private bool ValidateForm()
		{
			if (cboStep.EditValue == null)
			{
				MessageBox.Show("Xin hãy chọn một công đoạn.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (txtName.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy điền Mô tả chi tiết kiểm tra.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cboValueType.SelectedIndex < 0)
			{
				MessageBox.Show("Xin hãy chọn một kiểu kiểm tra.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			string sql1 = "";
			if (_KnifeJigWorking.ID > 0)
			{
				sql1 = $"SELECT top 1 ID from KnifeJigWorking WHERE SortOrder = '{txtSortOrder.Value}' and ID <>'{_KnifeJigWorking.ID}' and KnifeDetailListID={_KnifeJigWorking.KnifeDetailListID} and KnifeJigStepID={_KnifeJigWorking.KnifeStepID} ";
			}
			else
			{
				sql1 = $"SELECT top 1 ID from KnifeJigWorking WHERE SortOrder = '{txtSortOrder.Value}' and KnifeDetailListID={KnifeDetailListID} and KnifeJigStepID={StepID}";
			}
			int id = TextUtilsHP.ToInt(TextUtilsHP.ExcuteScalar(sql1));
			if (id > 0)
			{
				MessageBox.Show("Trùng STT", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//return false;
			}
			if (cboCheckValueType.SelectedIndex <= 0)
			{
				MessageBox.Show("Xin hãy chọn một loại giá trị kiểm tra.", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
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
			if (Global.UserID != 0)
			{
				_KnifeJigWorking.CreatedBy = Global.AppFullName;
				_KnifeJigWorking.UpdatedBy = Global.AppFullName;
			}
			_KnifeJigWorking.KnifeDetailListID = KnifeDetailListID;
			_KnifeJigWorking.KnifeStepID = TextUtilsHP.ToInt(cboStep.EditValue);
			_KnifeJigWorking.WorkingName = txtName.Text.Trim();
			_KnifeJigWorking.SortOrder = TextUtilsHP.ToInt(txtSortOrder.Value);
			_KnifeJigWorking.Unit = txtUnit.Text.Trim();
			_KnifeJigWorking.ValueType = cboValueType.SelectedIndex;// kiểu kiểm tra
			_KnifeJigWorking.ValueTypeName = cboValueType.SelectedIndex == 0 ? "Check mark" : "Giá trị\n数値";
			_KnifeJigWorking.MaxValue = TextUtilsHP.ToDecimal(txtMax.Value);
			_KnifeJigWorking.MinValue = TextUtilsHP.ToDecimal(txtMin.Value);
			_KnifeJigWorking.CheckValueType = cboCheckValueType.SelectedIndex;
			_KnifeJigWorking.ProductStepCode = cboStep.Text;
			_KnifeJigWorking.Frequency = txtFrequency.Text;
			if (_KnifeJigWorking.ValueType == 0)//checkmark
			{
				_KnifeJigWorking.PeriodValue = txtPeriodValue.Text.Trim();
			}
			else
			{
				if (_KnifeJigWorking.CheckValueType == 1)//giá trị dạng số
				{
					if (_KnifeJigWorking.MinValue == _KnifeJigWorking.MaxValue)
					{
						_KnifeJigWorking.PeriodValue = _KnifeJigWorking.MaxValue.ToString("n3");
					}
					else
					{
						_KnifeJigWorking.PeriodValue = _KnifeJigWorking.MinValue.ToString("n3") + "~" + _KnifeJigWorking.MaxValue.ToString("n3");
					}
				}
				else
				{
					_KnifeJigWorking.PeriodValue = txtPeriodValue.Text.Trim();
				}
			}

			if (_KnifeJigWorking.ID > 0)
			{
				KnifeJigWorkingBO.Instance.Update(_KnifeJigWorking);
				return true;
			}
			else
			{
				_KnifeJigWorking.ID = (int)KnifeJigWorkingBO.Instance.Insert(_KnifeJigWorking);
			}
			return true;
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
				_KnifeJigWorking = new KnifeJigWorkingModel();
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
	}
}
