using ST.Business;
using ST.Model;
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
	public partial class frmSequence : _Forms
	{
		public SequenceModel _SequenceModel = new SequenceModel();
		public frmSequence()
		{
			InitializeComponent();
		}
		/// <summary>
		/// load dữ liễu lên khi mở form
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void frmShelfDetail_Load(object sender, EventArgs e)
		{
			if (frmShelf.EditClick == 1)
			{
				btnSaveNew.Visible = false;
			}
			loadData();
			txtCode.Focus();
		}
		/// <summary>
		/// void loadData
		/// </summary>
		private void loadData()
		{
			txtCode.Text = _SequenceModel.SequenceCode;
			txtName.Text = _SequenceModel.SequenceName;
			if (txtCode.Text == "0")
			{
				txtCode.Clear();
			}
		}
		/// <summary>
		/// void lưu dữ liệu
		/// </summary>
		private void SaveGroup()
		{
			if (!ValidateForm()) return;
			_SequenceModel.SequenceName = txtName.Text;
			_SequenceModel.SequenceCode = txtCode.Text.Trim();
			if (_SequenceModel.ID > 0)
			{
				SequenceBO.Instance.Update(_SequenceModel);
			}
			else
			{
				SequenceBO.Instance.Insert(_SequenceModel);
			}

		}
		/// <summary>
		/// validate trước khi cất dữ liệu
		/// </summary>
		/// <returns></returns>
		bool ValidateForm()
		{
			if (txtCode.Text == "" || txtName.Text == "")
			{
				MessageBox.Show("Cần nhập đầy đủ thông tin vào các ô còn trống!", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			DataTable dt;
			if (_SequenceModel.ID > 0)
			{
				dt = TextUtilsStock.Select("select top 1 SequenceCode from Sequence where SequenceCode = '" + txtCode.Text.Trim() + "' and ID <> " + _SequenceModel.ID);
			}
			else
			{
				dt = TextUtilsStock.Select("select top 1 SequenceCode from Sequence where SequenceCode = '" + txtCode.Text.Trim() + "'");
			}
			if (dt != null)
			{
				if (dt.Rows.Count > 0)
				{
					MessageBox.Show("Tên nhóm đã tồn tại, xin hãy nhập tên khác!", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
					return false;
				}
			}
			return true;
		}

		/// <summary>
		/// event save dữ liệu _SequenceModel to database
		/// </summary>
		private void eventSaveGroup()
		{
			SaveGroup();
			this.DialogResult = DialogResult.OK;
			this.Close();
		}
		/// <summary>
		/// event save when used btnSave
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSave_Click(object sender, EventArgs e)
		{
			eventSaveGroup();
		}
		/// <summary>
		/// event save when used shortcut keys ctrl+S
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			eventSaveGroup();
		}
		/// <summary>
		/// save and creat new txtbox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			SaveGroup();
			txtCode.Clear();
			txtName.Clear();
		}

		private void frmShelfDetail_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void txtCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtName.Focus();
			}
		}
	}
}
