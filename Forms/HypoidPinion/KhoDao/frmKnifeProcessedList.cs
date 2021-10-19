using System;
using System.Collections;
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
	public partial class frmKnifeProcessedList : _Forms
	{
		#region Variables
		private KnifeProcessedListModel knifeProcessedList = new KnifeProcessedListModel();
		public int knifeID = 0;
		int prevRow;
		DataTable dtKnifeList;
		public string _KnifeCode = "";
		#endregion

		#region Methods
		public frmKnifeProcessedList()
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
			string sql = $"SELECT ID, KnifeCode, STD, ATC FROM dbo.KnifeDetailList WHERE Status = 1 and KnifeCode=N'{txtKnifeList.Text.Trim()}'";
			dtKnifeList = TextUtilsHP.Select(sql);
		}

		void LoadAllComboboxData()
		{
			LoadListUsers();
			LoadListKnifeList();

			DataTable dt = TextUtilsHP.Select(cGlobalFunction.GetWithSelectionQuery(new string[] { "ID", "MachineCode", "MachineName", "StageCode" }, "Machine", "MachineCode"));
			cMachine.Properties.DataSource = dt;
			cMachine.Properties.DisplayMember = "MachineCode";
			cMachine.Properties.ValueMember = "ID";
		}

		void LoadData()
		{
			string sql = "SELECT kdl.*, u.UserCode, u.DepartmentCode, Status = (CASE k.Status WHEN 1 THEN N'Còn sử dụng' WHEN 0 THEN N'Đã hủy' END) FROM dbo.KnifeProcessedList AS kdl, dbo.Users AS u, dbo.KnifeDetailList AS k WHERE kdl.WorkerID = u.ID AND kdl.KnifeID = k.ID";
			DataTable arr = TextUtilsHP.Select(sql);
			dtgvKnifeProcessedList.DataSource = arr;
		}

		void LoadDataToForm()
		{
			if (dtKnifeList.Rows.Count > 0)
				txtKnifeList.Text = TextUtilsHP.ToString(dtKnifeList.Rows[0]["KnifeCode"]);
		}

		void ClearFormData()
		{
			txbProductCode.Text = "";
			txbQuantity.Value = 0;
			txbProductCode_Leave(null, null);
			txbQuantity_Leave(null, null);
		}

		bool ValidateForm()
		{
			/*if (cKnifeList.EditValue == null)
			{
				MessageBox.Show("Vui lòng nhập mã dao!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}*/

			if (txbQuantity.Value == 0)
			{
				MessageBox.Show("Vui lòng nhập số lượng khác 0!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cWorker.EditValue == null)
			{
				MessageBox.Show("Vui lòng chọn người phụ trách!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}


			if (string.IsNullOrEmpty(txtKnifeList.Text))
			{
				MessageBox.Show("Vui lòng chọn loại dao!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			KnifeDetailListModel currentModel = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(TextUtilsHP.ToInt64(dtKnifeList.Rows[0]["ID"]));
			if (currentModel.CurrentATC == currentModel.ATC)
			{
				if (MessageBox.Show("Mã dao này không thể tiếp tục sử dụng! \n Bạn có muốn hủy mã dao này?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
				{
					frmKnifeDisposed frm = new frmKnifeDisposed();
					frm.knifeID = currentModel.ID;
					if (frm.ShowDialog() == DialogResult.OK)
					{
						knifeProcessedList = new KnifeProcessedListModel();
						ClearFormData();
						LoadData();
					}
				}
				return false;

			}

			if (currentModel.CurrentSTD == currentModel.STD)
			{
				if (MessageBox.Show("Mã dao này cần được mài trước khi sử dụng! \n Bạn có muốn mài mã dao này ngay bây giờ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
				{
					// Mai dao
					frmKnifeSharpen frm = new frmKnifeSharpen();
					frm.knifeID = TextUtilsHP.ToInt(dtKnifeList.Rows[0]["ID"]);
					if (frm.ShowDialog() == DialogResult.OK)
					{
						knifeProcessedList = new KnifeProcessedListModel();
						ClearFormData();
						LoadData();
					}
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
					knifeProcessedList.KnifeCode = txtKnifeList.Text.Trim();
					knifeProcessedList.KnifeID = TextUtilsHP.ToInt(dtKnifeList.Rows[0]["ID"]);
					knifeProcessedList.WorkerID = TextUtilsHP.ToInt(cWorker.EditValue);
					knifeProcessedList.ProductCode = txbProductCode.Text.Trim();
					knifeProcessedList.MachineID = TextUtilsHP.ToInt(cMachine.EditValue);
					knifeProcessedList.MachineCode = cMachine.Text.Trim();
					knifeProcessedList.Quantity = TextUtilsHP.ToInt(txbQuantity.Value);

					TextUtilsHP.ExcuteProcedure("spKnifeAddProcess",
									new string[] { "@knifeID", "@knifeCode", "@workerID", "@productCode", "@machineID", "@machineCode", "@quantity" },
									new object[] { knifeProcessedList.KnifeID, knifeProcessedList.KnifeCode, knifeProcessedList.WorkerID, knifeProcessedList.ProductCode, knifeProcessedList.MachineID, knifeProcessedList.MachineCode, knifeProcessedList.Quantity });
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
		private void frmKnifeSharpenList_Load(object sender, EventArgs e)
		{
			LoadAllComboboxData();
			LoadData();
			if (knifeID != 0)
			{
				txtKnifeList.Text = _KnifeCode;
				LoadDataToForm();
				cKnifeList_Leave(null, null);
			}
			txtKnifeList.Focus();
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		/*	private void timer1_Tick(object sender, EventArgs e)
			{
				txbTime.Text = string.Format("{0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString());
			}
	*/
		private void cWorker_EditValueChanged(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cWorker.Text))
			{
				DataRowView selectData = (DataRowView)cWorker.GetSelectedDataRow();
				string departmentCode = (string)selectData.Row.ItemArray[6];
				txbDepartmentCode.Text = departmentCode;
				txtKnifeList.Focus();
			}
		}

		private void frmKnifeSharpenList_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				knifeProcessedList = new KnifeProcessedListModel();
				ClearFormData();
				LoadData();
			}
		}

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

		private void txbProductCode_Leave(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(txbProductCode.Text))
			{
				txbProductCode.BackColor = Color.White;
			}
			else
			{
				txbProductCode.BackColor = Color.FromArgb(255, 153, 255);
			}
		}

		private void cMachine_Leave(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(cMachine.Text))
			{
				cMachine.BackColor = Color.White;
			}
			else
			{
				cMachine.BackColor = Color.FromArgb(255, 153, 255);
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

		private void txbQuantity_Leave(object sender, EventArgs e)
		{
			if (!string.Equals(txbQuantity.Text, "0"))
			{
				txbQuantity.BackColor = Color.White;
			}
			else
			{
				txbQuantity.BackColor = Color.FromArgb(255, 153, 255);
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

		private void gvKnifeProcessedList_DoubleClick(object sender, EventArgs e)
		{
			knifeID = TextUtilsHP.ToInt(gvKnifeProcessedList.GetFocusedRowCellValue(colKnifeID));
			if (knifeID == 0) return;
			LoadDataToForm();
		}
		#endregion

		private void cKnifeList_KeyPress(object sender, KeyPressEventArgs e)
		{
			cWorker.Focus();
		}

		private void cWorker_KeyPress(object sender, KeyPressEventArgs e)
		{
			txbProductCode.Focus();
		}

		private void cMachine_KeyPress(object sender, KeyPressEventArgs e)
		{
			txbQuantity.Focus();
		}

		private void cKnifeList_EditValueChanged(object sender, EventArgs e)
		{
			txbProductCode.Focus();
		}

		private void txbProductCode_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			cMachine.Focus();
			//LoadListKnifeList();
		}

		private void cMachine_EditValueChanged(object sender, EventArgs e)
		{
			txbQuantity.Focus();
		}

		private void txbQuantity_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtKnifeList.Focus();
		}

		private void txtKnifeList_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txbProductCode.Focus();
			LoadListKnifeList();
		}
	}
}
