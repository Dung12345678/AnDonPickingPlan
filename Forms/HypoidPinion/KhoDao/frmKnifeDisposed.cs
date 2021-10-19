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
	public partial class frmKnifeDisposed : _Forms
	{
		#region Variables
		private KnifeDisposedDetailsModel knifeDisposed = new KnifeDisposedDetailsModel();
		public int knifeID = 0;
		int prevRow;
		#endregion

		#region Methods
		public frmKnifeDisposed()
		{
			InitializeComponent();
		}

		void LoadListUsers()
		{
			DataTable dt = TextUtilsHP.Select(cGlobalFunction.GetWithSelectionQuery(new string[] { "ID", "UserCode", "FullName", "BirthOfDate", "DepartmentID", "DepartmentCode" }, "Users", "UserCode"));
			cWorker.Properties.DataSource = dt;
			cWorker.Properties.DisplayMember = "UserCode";
			cWorker.Properties.ValueMember = "ID";
		}

		void LoadListKnifeList()
		{
			string sql = "SELECT ID, KnifeCode, STD, ATC, CurrentSTD, CurrentATC FROM dbo.KnifeDetailList WHERE Status = 1";
			DataTable dt = TextUtilsHP.Select(sql);
			cKnifeList.Properties.DataSource = dt;
			cKnifeList.Properties.DisplayMember = "KnifeCode";
			cKnifeList.Properties.ValueMember = "ID";
		}

		void LoadData()
		{
			string sql = "SELECT kdl.*, u.UserCode, u.DepartmentCode FROM dbo.KnifeDisposedDetails AS kdl, dbo.Users AS u WHERE kdl.WorkerID = u.ID";
			DataTable arr = TextUtilsHP.Select(sql);
			dtgvKnifeDispose.DataSource = arr;
		}

		void LoadDataToForm()
		{
			cKnifeList.EditValue = knifeID;
		}

		void ClearFormData()
		{
			cKnifeList.EditValue = 0;
		}

		bool ValidateForm()
		{

			if (cWorker.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}


			if (string.IsNullOrEmpty(cKnifeList.Text))
			{
				MessageBox.Show("Vui lòng chọn loại dao!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			DataRowView selectData = (DataRowView)cKnifeList.GetSelectedDataRow();
			int atc = TextUtilsHP.ToInt(selectData.Row.ItemArray[3]);
			int currentATC = TextUtilsHP.ToInt(selectData.Row.ItemArray[5]);
			if (currentATC < atc)
			{
				if (MessageBox.Show("Mã dao này vẫn có thể tiếp tục sử dụng! \n Bạn vẫn muốn hủy mã dao này?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
				{
					return true;
				}
				return false;
			}
			return true;
		}

		bool SaveData()
		{
			if (ValidateForm())
			{
				try
				{
					knifeDisposed.KnifeCode = cKnifeList.Text.Trim();
					knifeDisposed.KnifeID = TextUtilsHP.ToInt(cKnifeList.EditValue);
					knifeDisposed.WorkerID = TextUtilsHP.ToInt(cWorker.EditValue);


					TextUtilsHP.ExcuteProcedure("spKnifeAddToDisposeQueue",
									new string[] { "@knifeCode", "@knifeID", "@workerID" },
									new object[] { knifeDisposed.KnifeCode, knifeDisposed.KnifeID, knifeDisposed.WorkerID });
					return true;
				}
				catch (Exception ex)
				{
					MessageBox.Show("Có lỗi trong quá trình xử lý!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			return false;
		}

		#endregion

		#region Events
	

		private void frmKnifeDisposed_Load(object sender, EventArgs e)
		{
			LoadListUsers();
			LoadListKnifeList();
			LoadData();
			LoadDataToForm();
		}

		private void frmKnifeDisposed_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cKnifeList_Leave(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cKnifeList.Text))
			{
				cKnifeList.BackColor = Color.White;
			}
			else
			{
				cKnifeList.BackColor = Color.FromArgb(255, 153, 255);
			}
		}

		private void cWorker_Leave(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cWorker.Text))
			{
				cWorker.BackColor = Color.White;
			}
			else
			{
				cWorker.BackColor = Color.FromArgb(255, 153, 255);
			}
		}

		private void cWorker_EditValueChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cWorker.Text))
			{
				DataRowView selectData = (DataRowView)cWorker.GetSelectedDataRow();
				string departmentCode = (string)selectData.Row.ItemArray[6];
				txbDepartmentCode.Text = departmentCode;
			}
		}

		private void txbDepartmentCode_TextChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(txbDepartmentCode.Text))
			{
				txbDepartmentCode.BackColor = Color.White;
			}
			else
			{
				txbDepartmentCode.BackColor = Color.FromArgb(255, 153, 255);
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void gvKnifeDispose_DoubleClick(object sender, EventArgs e)
		{
			knifeID = TextUtilsHP.ToInt(gvKnifeDispose.GetFocusedRowCellValue(colKnifeID));
			if (knifeID == 0) return;
			LoadDataToForm();
		}

		private void cKnifeList_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				cWorker.Focus();
			}
		}

		private void cWorker_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				cKnifeList.Focus();
			}
		}

		private void txbDepartmentCode_KeyPress(object sender, KeyPressEventArgs e)
		{
			//cKnifeList.Focus();
		}

		#endregion


	}
}
