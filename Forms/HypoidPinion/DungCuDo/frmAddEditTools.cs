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
	public partial class frmAddEditTools : _Forms
	{
		#region Variables
		public ToolsModel _toolsModel = new ToolsModel();
		public int KnifeGroupID;
		public int Loai;
		#endregion

		#region Methods
		public frmAddEditTools()
		{
			InitializeComponent();
		}
		private void frmAddEditJig_Load(object sender, EventArgs e)
		{
			LoadDataToForm();
		}

		void LoadDataToForm()
		{
			txtCode.Text = _toolsModel.Code;
			txtName.Text = _toolsModel.Name;
			NUDMin.Value = TextUtilsHP.ToDecimal(_toolsModel.Min);
			NUDMax.Value = TextUtilsHP.ToDecimal(_toolsModel.Max);
			txtStd.Text = TextUtilsHP.ToString(_toolsModel.Std);
		}

		bool ValidateForm()
		{
			if (txtCode.Text == "")
			{
				MessageBox.Show("Vui lòng nhập mã dụng cụ đo!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			DataTable dt;
			if (_toolsModel.ID > 0)
			{
				dt = TextUtilsHP.Select("select 1 from Tools where Code = '" + txtCode.Text.Trim() + "' and ID <> " + _toolsModel.ID);
			}
			else
			{
				dt = TextUtilsHP.Select("select 1 from Tools where Code = '" + txtCode.Text.Trim() + "'");
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
					_toolsModel.Code = txtCode.Text.Trim();
					_toolsModel.Name = txtName.Text.Trim();
					_toolsModel.Std = TextUtilsHP.ToDouble(txtStd.Text.Trim());
					_toolsModel.Min = NUDMin.Value;
					_toolsModel.Max = NUDMax.Value;
					_toolsModel.NameDisplay = _toolsModel.Name + " " + _toolsModel.Min + " - " + _toolsModel.Max;
					_toolsModel.ConditionDisplay = _toolsModel.Min + " - " + _toolsModel.Max;
					_toolsModel.CreateDate = DateTime.Now;
					if (_toolsModel.ID > 0)
					{
						ToolsBO.Instance.Update(_toolsModel);
						return true;
					}
					else
					{
						_toolsModel.GroupID = KnifeGroupID;
						_toolsModel.ID = (int)ToolsBO.Instance.Insert(_toolsModel);
						TextUtilsHP.ExcuteSQL(string.Format("EXEC dbo.CreateToolsWorking_ByGroupID {0},{1}", _toolsModel.GroupID, _toolsModel.ID));
						return true;
					}
					//if (_toolsModel.ID > 0)
					//{
					//	ToolsBO.Instance.Update(_toolsModel);
					//	return true;
					//}
					//else
					//{
					//	_toolsModel.GroupID = KnifeGroupID;
					//	ToolsBO.Instance.Insert(_toolsModel);

					//	return true;
					//}

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
				_toolsModel = new ToolsModel();
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

		private void txtJigName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtStd.Focus();
		}
		private void txtCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtName.Focus();
		}

		private void txtStd_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			NUDMin.Focus();

		}

		private void txtMin_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			NUDMax.Focus();
		}
	}
}
