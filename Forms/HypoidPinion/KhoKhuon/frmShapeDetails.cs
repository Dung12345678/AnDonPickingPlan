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
	public partial class frmShapeDetails : _Forms
	{
		#region Variables
		public KnifeDetailListModel _KnifeDetailListModel = new KnifeDetailListModel();
		public int KnifeGroupID;
		public int Loai;
		#endregion

		#region Methods
		public frmShapeDetails()
		{
			InitializeComponent();
		}
		private void frmAddEditJig_Load(object sender, EventArgs e)
		{
			LoadListUsers();
			LoadJigCha();
			LoadCboGroup();
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
			DataTable dt = TextUtilsHP.Select("Select * from KnifeDetailList where Type=2 and ParentID=0");
			cbShapeParent.DataSource = dt;
			cbShapeParent.DisplayMember = "KnifeCode";
			cbShapeParent.ValueMember = "ID";
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
			txtShapeCode.Text = _KnifeDetailListModel.KnifeCode;
			cWorker.EditValue = _KnifeDetailListModel.WorkerID;
			txtShapeName.Text = _KnifeDetailListModel.KnifeName;
			if (KnifeGroupID == 0)
			{
				cboGroupKnife.SelectedValue = _KnifeDetailListModel.KnifeGroupID;
			}
			else
				cboGroupKnife.SelectedValue = KnifeGroupID;
			txtDescription.Text = _KnifeDetailListModel.Description;
			cbShapeParent.SelectedValue = _KnifeDetailListModel.ParentID;
			//txtQty.Value = _KnifeDetailListModel.Qty;
		}

		bool ValidateForm()
		{
			if (txtShapeCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã Jig!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
				dt = TextUtilsHP.Select($"select 1 from KnifeDetailList where KnifeCode = '{txtShapeCode.Text.Trim()}' and ID <> {_KnifeDetailListModel.ID}  and KnifeGroupID = '{TextUtilsHP.ToInt(cboGroupKnife.SelectedValue)}'");
			}
			else
			{
				dt = TextUtilsHP.Select($"select 1 from KnifeDetailList where KnifeCode = '{txtShapeCode.Text.Trim()}' and KnifeGroupID = '{TextUtilsHP.ToInt(cboGroupKnife.SelectedValue)}' ");
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
					_KnifeDetailListModel.KnifeCode = txtShapeCode.Text.Trim();
					_KnifeDetailListModel.WorkerID = Lib.ToInt(cWorker.EditValue);
					_KnifeDetailListModel.KnifeName = txtShapeName.Text.Trim();
					_KnifeDetailListModel.ImportedDate = DateTime.Now;
					_KnifeDetailListModel.KnifeGroupID = TextUtilsHP.ToInt(cboGroupKnife.SelectedValue);
					//_KnifeDetailListModel.Qty = TextUtilsHP.ToInt(txtQty.Value);
					_KnifeDetailListModel.ParentID = Lib.ToInt(cbShapeParent.SelectedValue);
					_KnifeDetailListModel.Description = txtDescription.Text;
					_KnifeDetailListModel.Type = 2; // 0: Dao, 1: Jig, 2 :Khuôn

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
			txtDescription.Focus();
		}

		private void txtJigCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtShapeName.Focus();

		}

		private void txtJigName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			cbShapeParent.Focus();
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
			//if (e.KeyCode != Keys.Enter) return;
			//txtQty.Focus();
		}
	}
}
