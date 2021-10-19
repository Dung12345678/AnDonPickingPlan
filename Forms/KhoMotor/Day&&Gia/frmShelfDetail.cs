using ST.Business;
using ST.Model;
using ExcelDataReader;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace BMS
{
	public partial class frmShelfDetail : _Forms
	{
		public ShelfModel oShelfModel = new ShelfModel();

		public frmShelfDetail()
		{
			InitializeComponent();
		}
		/// <summary>
		/// load dữ liệu lên khi load form
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void frmShelfDetail_Load(object sender, EventArgs e)
		{
			loadGroup();
			loadData();
		}
		private void loadGroup()
		{
			cboGroupnhom.Properties.DisplayMember = "SequenceName";
			cboGroupnhom.Properties.ValueMember = "ID";
			cboGroupnhom.Properties.DataSource = SequenceBO.Instance.FindAll();
		}
		private void loadData()
		{
			cboGroupnhom.EditValue = oShelfModel.SequenceID;
			txtTenGiaDo.Text = oShelfModel.ShelfCode;
			txtGiaDo.Text = oShelfModel.ShelfName;
			txtSTT.Text = oShelfModel.SortOrder.ToString();
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			SaveData();
			this.DialogResult = DialogResult.OK;
		}
		void SaveData()
		{
			if (ValidateForm())
			{
				try
				{
					oShelfModel.SequenceID = TextUtilsStock.ToInt(cboGroupnhom.EditValue);
					oShelfModel.ShelfCode = txtGiaDo.Text.Trim();
					oShelfModel.ShelfName = txtTenGiaDo.Text.Trim();
					oShelfModel.SortOrder = TextUtilsStock.ToInt(txtSTT.Text.Trim());
					oShelfModel.CreatDate = DateTime.Now;

					if (oShelfModel.ID > 0)
					{
						ShelfBO.Instance.Update(oShelfModel);
					}
					else
					{
						ShelfBO.Instance.Insert(oShelfModel);
					}
				}
				catch (Exception)
				{
					MessageBox.Show("Lỗi Update sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			SaveData();
			txtTenGiaDo.Clear();
			txtGiaDo.Clear();
			txtSTT.Clear();
		}

		private void frmShelfDetail_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			SaveData();
			this.DialogResult = DialogResult.OK;
		}
		/// <summary>
		/// check lỗi
		/// </summary>
		/// <returns></returns>
		/// 

		bool ValidateForm()
		{
			if (cboGroupnhom.Text == "" || txtTenGiaDo.Text == "" || txtGiaDo.Text == "" || txtSTT.Text == "")
			{
				MessageBox.Show("Cần nhập đầy đủ thông tin vào các ô còn trống!", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			DataTable dt;
			if (oShelfModel.ID > 0)
			{
				dt = TextUtilsStock.Select($"select top 1 ShelfCode from Shelf where ShelfCode = '{txtGiaDo.Text.Trim()}' and SequenceID='{cboGroupnhom.EditValue}' and ID <> " + oShelfModel.ID);
			}
			else
			{
				dt = TextUtilsStock.Select($"select top 1 ShelfCode from Shelf where ShelfCode = '{txtGiaDo.Text.Trim()}' and SequenceID='{cboGroupnhom.EditValue}'");
			}
			if (dt != null)
			{
				if (dt.Rows.Count > 0)
				{
					MessageBox.Show("Mã giá đỡ này đã tồn tại trong danh sách! Bạn muốn thêm mới giá đỗ có cùng mã ?", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
					return false;
				}
			}
			return true;
		}

		private void txtSTT_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtGiaDo.Focus();
			}
		}

		private void txtGiaDo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtTenGiaDo.Focus();
			}
		}

		private void txtTenGiaDo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
			}
		}

		private void dtpCreateDate_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				btnSave.PerformClick();
			}
		}

		private void SaveNewToolStripMenuItem_Click(object sender, EventArgs e)
		{

		}
	}
}
