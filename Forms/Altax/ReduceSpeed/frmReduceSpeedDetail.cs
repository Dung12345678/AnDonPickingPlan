using BMS.Business;
using BMS.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmReduceSpeedDetail : _Forms
	{
		public SpeedReducerModel _SpeedReducerModel = new SpeedReducerModel();
		public frmReduceSpeedDetail()
		{
			InitializeComponent();
		}
		private void ProductionPlanDetail_Load(object sender, EventArgs e)
		{
			loadData();
			txtName.Focus();
		}
		// gán từ form frmProceductionPlan sang Detail
		// lấy dữ liệu từ form 1 sang form 2 
		//  public void funData(TextBox txtForm1) { lb.Text = txtForm1.Text; }
		// kiểm tra trước khi cất
		private bool ValidateForm()
		{
			if (txtName.Text == "")
			{
				MessageBox.Show("Xin hãy nhập tên điều kiện", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			else
			{
				DataTable dt;
				if (_SpeedReducerModel.ID > 0)
				{
					dt = TextUtils.Select("select top 1 1 from SpeedReducer where Name = '" + txtName.Text.Trim() + "' and ID <> " + _SpeedReducerModel.ID);
				}
				else
				{
					dt = TextUtils.Select("select top 1 1 from SpeedReducer where Name = '" + txtName.Text.Trim() + "'");
				}
				if (dt != null)
				{
					if (dt.Rows.Count > 0)
					{
						MessageBox.Show("Mã này đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}
			}
			if (txtStandValue.Text == "")
			{
				MessageBox.Show("Xin hãy điền giá trị tiêu chuẩn", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (NLine.Text == "")
			{
				MessageBox.Show("Xin hãy điền tên line", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}

		bool saveData()
		{
			if (!ValidateForm()) return false;
			_SpeedReducerModel.Name = txtName.Text.Trim();
			_SpeedReducerModel.Line = TextUtils.ToInt(NLine.Value);
			_SpeedReducerModel.StandValue = txtStandValue.Text.Trim(); ;
			_SpeedReducerModel.Min = NMin.Value;
			_SpeedReducerModel.Max = NMax.Value;
			_SpeedReducerModel.TimeGetVongQuay = TextUtils.ToDecimal(NTimeGetDataVongQuay.Value);
			_SpeedReducerModel.TimeRunF = TextUtils.ToDecimal(NTimeRunF.Value);
			_SpeedReducerModel.TimeTeaching = TextUtils.ToDecimal(NudTimeTeaching.Value);
			if (_SpeedReducerModel.ID > 0)
			{
				SpeedReducerBO.Instance.Update(_SpeedReducerModel);
				MessageBox.Show("Update thành công").ToString();
			}
			else
			{
				_SpeedReducerModel.CreateDate = DateTime.Now;
				SpeedReducerBO.Instance.Insert(_SpeedReducerModel);
				MessageBox.Show("Thêm mới thành công").ToString();
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
				_SpeedReducerModel = new SpeedReducerModel();
				loadData();
			}
		}
		void loadData()
		{
			txtName.Text = _SpeedReducerModel.Name;
			txtStandValue.Text = _SpeedReducerModel.StandValue;
			NLine.Value = _SpeedReducerModel.Line;
			NMin.Value = _SpeedReducerModel.Min;
			NMax.Value = _SpeedReducerModel.Max;
			NTimeGetDataVongQuay.Value = _SpeedReducerModel.TimeGetVongQuay;
			NTimeRunF.Value= _SpeedReducerModel.TimeRunF;
			NudTimeTeaching.Value= _SpeedReducerModel.TimeTeaching;
		}

		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSave_Click(null, null);
		}

		private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void lb_Click(object sender, EventArgs e)
		{
			//string sql = "select * from ProductionPlan where ID=" +lb.Text;
			//sql.
		}

		private void closeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void NMin_ValueChanged(object sender, EventArgs e)
		{
			txtStandValue.Text = NMin.Value + "~" + NMax.Value;
		}

		private void NMax_ValueChanged(object sender, EventArgs e)
		{
			txtStandValue.Text = NMin.Value + "~" + NMax.Value;
		}
	}
}
