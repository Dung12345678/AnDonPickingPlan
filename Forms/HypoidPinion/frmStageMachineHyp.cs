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
	public partial class frmStageMachineHyp : _Forms
	{
		#region Variables
		int prevRow;
		int prevRow1;
		bool showAll = false;
		#endregion

		#region Method
		public frmStageMachineHyp()
		{
			InitializeComponent();
		}
		void LoadStage()
		{
			DataTable arr = TextUtils.Select(cGlobalFunction.GetAllQuery("Stage", "StageCode"));
			dtgvStage.DataSource = arr;
		}
		void LoadMachine()
		{
			if (cbShowAll.Checked)
				showAll = true;
			else
				showAll = false;
			int id = TextUtils.ToInt(gvStage.GetFocusedRowCellValue(colIDStage));
			if (id == 0) return;
			DataTable arr = TextUtils.LoadDataFromSP("spGetMachineListWithFilter", "a", new string[] { "@keyword", "@stageID", "@showAll" }, new object[] { txbSearch.Text.Trim(), id, showAll });
			dtgvMachine.DataSource = arr;
		}
		#endregion

		#region Event
		private void btnCreateStage_Click(object sender, EventArgs e)
		{
			frmAddEditStageHyp frm = new frmAddEditStageHyp(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadStage();
				LoadMachine();
			}
		}

		private void btnEditStage_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(gvStage.GetFocusedRowCellValue(colIDStage));
			if (id == 0) return;
			prevRow = gvStage.GetSelectedRows()[0];
			StageModel model = (StageModel)StageBO.Instance.FindByPK(id);

			frmAddEditStageHyp frm = new frmAddEditStageHyp(cGlobalVariables.Edit);
			frm.stageModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadStage();
				LoadMachine();
				gvStage.FocusedRowHandle = prevRow;
			}
		}

		private void btnDelStage_Click(object sender, EventArgs e)
		{
			if (!gvStage.IsDataRow(gvStage.FocusedRowHandle))
				return;
			int strID = TextUtils.ToInt(gvStage.GetFocusedRowCellValue(colIDStage));
			string str = TextUtils.ToString(gvStage.GetFocusedRowCellValue("StageCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa công đoạn [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					if (MachineBO.Instance.CheckExist("StageID", strID))
					{
						MessageBox.Show("Công đoạn này đã tồn tại các máy trực thuộc!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					else
					{
						StageBO.Instance.Delete(strID);
						gvStage.DeleteRow(gvStage.FocusedRowHandle);
						MessageBox.Show("Xoá thành công", TextUtils.Caption);
					}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void btnRefresh_Click(object sender, EventArgs e)
		{
			LoadStage();
			LoadMachine();
		}

		private void btnCreateMachine_Click(object sender, EventArgs e)
		{
			frmAddEditMachineHyp frm = new frmAddEditMachineHyp(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadMachine();
			}
		}

		private void btnEditMachine_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(gvMachine.GetFocusedRowCellValue(colIDMac));
			if (id == 0) return;
			prevRow = gvMachine.GetSelectedRows()[0];
			MachineModel model = (MachineModel)MachineBO.Instance.FindByPK(id);

			frmAddEditMachineHyp frm = new frmAddEditMachineHyp(cGlobalVariables.Edit);
			frm.machineModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadMachine();
				gvMachine.FocusedRowHandle = prevRow;
			}
		}

		private void btnDelMachine_Click(object sender, EventArgs e)
		{
			if (!gvMachine.IsDataRow(gvMachine.FocusedRowHandle))
				return;
			int strID = TextUtils.ToInt(gvMachine.GetFocusedRowCellValue(colIDMac));
			string str = TextUtils.ToString(gvMachine.GetFocusedRowCellValue("MachineCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa máy [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					MachineBO.Instance.Delete(strID);
					gvMachine.DeleteRow(gvMachine.FocusedRowHandle);
					MessageBox.Show("Xoá thành công", TextUtils.Caption);
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void btnShowWorker_Click(object sender, EventArgs e)
		{
			prevRow = gvMachine.GetSelectedRows()[0];
			frmStageWorkerHyp frm = new frmStageWorkerHyp();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadMachine();
				gvMachine.FocusedRowHandle = prevRow;
			}
		}
		private void gvStage_DoubleClick(object sender, EventArgs e)
		{
			btnEditStage_Click(null, null);
		}

		private void gvStage_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
		{
			LoadMachine();
		}

		private void gvMachine_DoubleClick(object sender, EventArgs e)
		{
			btnEditMachine_Click(null, null);
		}

		private void frmStageMachineHyp_Load(object sender, EventArgs e)
		{
			LoadStage();
			LoadMachine();
		}

		private void btnSearch_Click(object sender, EventArgs e)
		{
			LoadMachine();
		}

		private void btnCancelSearch_Click(object sender, EventArgs e)
		{
			txbSearch.Text = "";
			LoadMachine();
		}

		private void cbShowAll_CheckedChanged(object sender, EventArgs e)
		{
			LoadMachine();
		}
		private void txbSearch_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
				btnSearch_Click(null, null);
		}
		#endregion

		
	}
}
