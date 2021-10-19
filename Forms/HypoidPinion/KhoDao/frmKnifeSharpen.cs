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
	public partial class frmKnifeSharpen : _Forms
	{
		#region Variables
		private KnifeSharpeningDetailsModel knifeSharpeningModel = new KnifeSharpeningDetailsModel();
		public int knifeID = 0;
		DataTable dtKnifeDetailList;
		int prevRow;
		#endregion

		#region Methods
		public frmKnifeSharpen()
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
			string sql = $"SELECT ID, KnifeCode, STD, ATC, CurrentSTD, CurrentATC FROM dbo.KnifeDetailList WHERE Status = 1 and KnifeCode=N'{txtKnifeList.Text.Trim()}'";
			dtKnifeDetailList = TextUtilsHP.Select(sql);
			if (dtKnifeDetailList.Rows.Count <= 0)
			{
				MessageBox.Show("Không tìm thấy mã Dao", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				txtKnifeList.Focus();
				txtKnifeList.SelectAll();
				return;
			}

		}

		void LoadData()
		{
			string sql = "SELECT kdl.*, u.UserCode, u.DepartmentCode, Type = (CASE k.Status WHEN 1 THEN N'Còn sử dụng' WHEN 0 THEN N'Đã hủy' END) FROM dbo.KnifeSharpeningDetails AS kdl, dbo.Users AS u, dbo.KnifeDetailList AS k WHERE kdl.WorkerID = u.ID AND kdl.KnifeID = k.ID";
			//string sql = "SELECT kdl.*, u.UserCode, u.DepartmentCode FROM dbo.KnifeSharpeningDetails AS kdl, dbo.Users AS u WHERE kdl.WorkerID = u.ID";
			DataTable arr = TextUtilsHP.Select(sql);
			dtgvKnifeSharpen.DataSource = arr;
		}

		void LoadDataToForm()
		{
			if (dtKnifeDetailList.Rows.Count > 0)
				txtKnifeList.Text = TextUtilsHP.ToString(dtKnifeDetailList.Rows[0]["KnifeCode"]);
		}

		void ClearFormData()
		{
			txtKnifeList.Text = "";
			txbSTD.Text = "";
			txbATC.Text = "";
		}

		bool ValidateForm()
		{

			if (cWorker.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (string.IsNullOrEmpty(txtKnifeList.Text))
			{
				MessageBox.Show("Vui lòng nhập mã Dao!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (dtKnifeDetailList.Rows.Count <= 0)
			{
				MessageBox.Show("Không tìm thấy mã Dao vui lòng nhập lại!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			//DataRowView selectData = (DataRowView)txtKnifeList.GetSelectedDataRow();
			int std = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["STD"]);
			int atc = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["ATC"]);
			int currentSTD = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["CurrentSTD"]);
			int currentATC = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["CurrentATC"]);
			if (currentATC == atc)
			{
				if (MessageBox.Show("Mã dao này không thể tiếp tục sử dụng! \n Bạn có muốn hủy mã dao này?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
				{
					// Huy dao
					frmKnifeDisposed frm = new frmKnifeDisposed();
					frm.knifeID = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["ID"]);
					if (frm.ShowDialog() == DialogResult.OK)
					{
						knifeSharpeningModel = new KnifeSharpeningDetailsModel();
						ClearFormData();
						LoadData();
					}
				}
				return false;


			}

			/*if (currentModel.CurrentSTD == currentModel.STD)
			{
				if (MessageBox.Show("Mã dao này cần được mài trước khi sử dụng! \n Bạn có muốn mài mã dao này ngay bây giờ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
				{
					// Mai dao
					frmKnifeSharpen frm = new frmKnifeSharpen();
					frm.knifeID = TextUtilsHP.ToInt(cKnifeList.EditValue);
					if (frm.ShowDialog() == DialogResult.OK)
					{

					}
				}
				return false;

			}*/
			return true;
		}

		bool SaveData()
		{
			if (ValidateForm())
			{
				try
				{
					knifeSharpeningModel.KnifeCode = txtKnifeList.Text.Trim();
					knifeSharpeningModel.KnifeID = TextUtilsHP.ToInt(dtKnifeDetailList.Rows[0]["ID"]);
					knifeSharpeningModel.WorkerID = TextUtilsHP.ToInt(cWorker.EditValue);


					TextUtilsHP.ExcuteProcedure("spKnifeAddToSharpenQueue",
									new string[] { "@knifeCode", "@knifeID", "@workerID" },
									new object[] { knifeSharpeningModel.KnifeCode, knifeSharpeningModel.KnifeID, knifeSharpeningModel.WorkerID });
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
		private void cKnifeList_Leave(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(txtKnifeList.Text))
			{
				txtKnifeList.BackColor = Color.White;
			}
			else
			{
				txtKnifeList.BackColor = Color.FromArgb(255, 153, 255);
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

		private void txbSTD_TextChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(txbSTD.Text))
			{
				txbSTD.BackColor = Color.White;
			}
			else
			{
				txbSTD.BackColor = Color.FromArgb(255, 153, 255);
			}
		}

		private void txbATC_TextChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(txbATC.Text))
			{
				txbATC.BackColor = Color.White;
			}
			else
			{
				txbATC.BackColor = Color.FromArgb(255, 153, 255);
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

		private void timer1_Tick(object sender, EventArgs e)
		{
			txbTime.Text = string.Format("{0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString());
		}

		private void frmKnifeSharpen_Load(object sender, EventArgs e)
		{
			txbTime.Text = TextUtils.ToString(DateTime.Now.ToString("dd-MM-yyyy"));
			LoadListUsers();
			//LoadListKnifeList();
			LoadData();
			//LoadDataToForm();
		}

		private void frmKnifeSharpen_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cWorker_EditValueChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cWorker.Text))
			{
				DataRowView selectData = (DataRowView)cWorker.GetSelectedDataRow();
				string departmentCode = (string)selectData.Row.ItemArray[6];
				txbDepartmentCode.Text = departmentCode;
			}
			txtKnifeList.Focus();
		}

		private void gvKnifeSharpen_DoubleClick(object sender, EventArgs e)
		{
			knifeID = TextUtilsHP.ToInt(gvKnifeSharpen.GetFocusedRowCellValue(colKnifeID));
			if (knifeID == 0) return;
			LoadDataToForm();
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				knifeSharpeningModel = new KnifeSharpeningDetailsModel();
				ClearFormData();
				//LoadListKnifeList();
				LoadData();
				txtKnifeList.Focus();
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
		#endregion

		private void cWorker_KeyPress(object sender, KeyPressEventArgs e)
		{
			txtKnifeList.Focus();
		}

		private void cKnifeList_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			LoadListKnifeList();

			if (!string.IsNullOrEmpty(txtKnifeList.Text))
			{
				if (dtKnifeDetailList.Rows.Count <= 0) return;
				//KnifeDetailListModel currentModel = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(TextUtilsHP.ToInt64(cKnifeList.EditValue));
				//DataRowView selectData = (DataRowView)txtKnifeList.GetSelectedDataRow();
				//string std = TextUtilsHP.ToString(selectData.Row.ItemArray[2]);
				string std = TextUtilsHP.ToString(dtKnifeDetailList.Rows[0]["STD"]);
				string atc = TextUtilsHP.ToString(dtKnifeDetailList.Rows[0]["ATC"]);
				string currentSTD = TextUtilsHP.ToString(dtKnifeDetailList.Rows[0]["CurrentSTD"]);
				string currentATC = TextUtilsHP.ToString(dtKnifeDetailList.Rows[0]["CurrentATC"]);
				txbSTD.Text = string.Format("{0} / {1}", currentSTD, std);
				txbATC.Text = string.Format("{0} / {1}", currentATC, atc);
				txbSTD.Focus();
			}
			cKnifeList_Leave(null, null);
		}
	}
}
