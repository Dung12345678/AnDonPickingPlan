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
	public partial class frmAddEditJig : _Forms
	{
		#region Variables
		public KnifeDetailListModel _KnifeDetailListModel = new KnifeDetailListModel();
		public int KnifeGroupID;
		public int Loai;
		#endregion

		#region Methods
		public frmAddEditJig()
		{
			InitializeComponent();
		}
		private void frmAddEditJig_Load(object sender, EventArgs e)
		{
			LoadCboGroup();
			LoadListUsers();
			LoadJigCha();
			LoadDataToForm();
		}
		void LoadCboGroup()
		{
			DataTable dt = TextUtilsHP.Select($"SELECT * FROM [HypoidPinion].[dbo].[KnifeGroup] Where Type={Loai}");
			cboGroupKnife.DataSource = dt;
			cboGroupKnife.DisplayMember = "KnifeGroupCode";
			cboGroupKnife.ValueMember = "ID";
		}
		void LoadJigCha()
		{
			DataTable dt = TextUtilsHP.Select("Select * from KnifeDetailList where Type=1 and ParentID=0");
			cbJigCha.DataSource = dt;
			cbJigCha.DisplayMember = "KnifeCode";
			cbJigCha.ValueMember = "ID";
		}
		void LoadListUsers()
		{
			DataTable dt = TextUtilsHP.Select(cGlobalFunction.GetWithSelectionQuery(new string[] { "ID", "UserCode", "FullName", "BirthOfDate", "DepartmentID", "DepartmentCode" }, "Users", "UserCode"));
			cWorker.Properties.DataSource = dt;
			cWorker.Properties.DisplayMember = "UserCode";
			cWorker.Properties.ValueMember = "ID";
		}
		void LoadDataToForm()
		{
			txtJigCode.Text = _KnifeDetailListModel.KnifeCode;
			cWorker.EditValue = _KnifeDetailListModel.WorkerID;
			txtJigName.Text = _KnifeDetailListModel.KnifeName;
			if (KnifeGroupID == 0)
			{
				cboGroupKnife.SelectedValue = _KnifeDetailListModel.KnifeGroupID;
			}
			else
				cboGroupKnife.SelectedValue = KnifeGroupID;
			txtDescription.Text = _KnifeDetailListModel.Description;
			cbJigCha.SelectedValue = _KnifeDetailListModel.ParentID;
			txtQty.Value = _KnifeDetailListModel.Qty;
			txtDateOld.Text = TextUtilsHP.ToString(_KnifeDetailListModel.DateOld);
			txtTanSuatKTra.Text = TextUtilsHP.ToString(_KnifeDetailListModel.FrequencyCheck);
		}

		bool ValidateForm()
		{
			if (txtJigCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã Jig!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}	
			if (txtTanSuatKTra.Text == "")
			{
				MessageBox.Show("Vui lòng nhập tần suất kiểm tra Jig!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			//if (txtWorkerName.Text == "")
			//{
			//	MessageBox.Show("Vui lòng nhập tên!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//	return false;
			//}
			DataTable dt;
			if (_KnifeDetailListModel.ID > 0)
			{
				dt = TextUtilsHP.Select($"select 1 from KnifeDetailList where KnifeCode = '{txtJigCode.Text.Trim()}' and ID <> {_KnifeDetailListModel.ID} and KnifeGroupID={TextUtilsHP.ToInt(cboGroupKnife.SelectedValue)}");
			}
			else
			{
				dt = TextUtilsHP.Select($"select 1 from KnifeDetailList where KnifeCode = '{txtJigCode.Text.Trim()}' and KnifeGroupID={TextUtilsHP.ToInt(cboGroupKnife.SelectedValue)}");
			}
			if (dt != null)
			{
				if (dt.Rows.Count > 0)
				{
					MessageBox.Show("Mã này đã tồn tại!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
					_KnifeDetailListModel.KnifeCode = txtJigCode.Text.Trim();
					_KnifeDetailListModel.WorkerID = Lib.ToInt(cWorker.EditValue);
					_KnifeDetailListModel.KnifeName = txtJigName.Text.Trim();
					_KnifeDetailListModel.ImportedDate = DateTime.Now;
					_KnifeDetailListModel.Qty = TextUtilsHP.ToInt(txtQty.Value);
					_KnifeDetailListModel.ParentID = Lib.ToInt(cbJigCha.SelectedValue);
					_KnifeDetailListModel.Description = txtDescription.Text;
					_KnifeDetailListModel.Type = 1;
					_KnifeDetailListModel.KnifeGroupID = TextUtils.ToInt(cboGroupKnife.SelectedValue);
					_KnifeDetailListModel.DateOld = TextUtilsHP.ToDate2(txtDateOld.Text.Trim());
					_KnifeDetailListModel.FrequencyCheck = txtTanSuatKTra.Text.Trim();
					if (_KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "TUẦN" || _KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "WEEK")
					{
						_KnifeDetailListModel.TypeFrequencyCheck = 0;
						_KnifeDetailListModel.ValueFrequency = 7;
					}
					else if (_KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "NGÀY" || _KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "DAILY" || _KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "DAY")
					{
						_KnifeDetailListModel.TypeFrequencyCheck = 1;
						_KnifeDetailListModel.ValueFrequency = 1;
					}
					else if (_KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "THÁNG" || _KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "MONTH")
					{
						_KnifeDetailListModel.TypeFrequencyCheck = 2;
						_KnifeDetailListModel.ValueFrequency = 1;
					}
					else if (_KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "NĂM" || _KnifeDetailListModel.FrequencyCheck.Trim().ToUpper() == "YEAR")
					{
						_KnifeDetailListModel.TypeFrequencyCheck = 3;
						_KnifeDetailListModel.ValueFrequency = 1;
					}
					if (_KnifeDetailListModel.ID > 0)
					{
						KnifeDetailListBO.Instance.Update(_KnifeDetailListModel);
						return true;
					}
					else
					{
						_KnifeDetailListModel.ID = (int)KnifeDetailListBO.Instance.Insert(_KnifeDetailListModel);
						TextUtilsHP.ExcuteSQL(string.Format("EXEC dbo.CreateKnifeJigWorking_ByGroupID {0},{1}", _KnifeDetailListModel.KnifeGroupID, _KnifeDetailListModel.ID));
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
				_KnifeDetailListModel = new KnifeDetailListModel();
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
		private void frmAddEditKnife_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
		#endregion

		private void cWorker_EditValueChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cWorker.Text))
			{
				DataRowView selectData = (DataRowView)cWorker.GetSelectedDataRow();
				string departmentCode = (string)selectData.Row.ItemArray[6];
				txtPart.Text = departmentCode;
			}

		}

		private void txtJigCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtJigName.Focus();

		}

		private void txtJigName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			cbJigCha.Focus();
		}

		private void cbJigCha_SelectedIndexChanged(object sender, EventArgs e)
		{
			cWorker.Focus();
		}

		private void txtPart_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtDescription.Focus();
		}

		private void txtDescription_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtQty.Focus();

		}
	}
}
