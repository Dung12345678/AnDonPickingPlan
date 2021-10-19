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
	public partial class frmStageWorkerHyp : _Forms
	{
		#region Variables
		int prevRow;
		#endregion

		#region Methods
		public frmStageWorkerHyp()
		{
			InitializeComponent();
			LoadData();
		}

		void LoadData()
		{
			DataTable arr = TextUtils.Select(cGlobalFunction.GetAllQuery("StageWorker", "StageCode"));
			dtgvStageWorker.DataSource = arr;
		}
		#endregion

		#region Events
		private void btnCreate_Click(object sender, EventArgs e)
		{
			frmAddEditStageWorker frm = new frmAddEditStageWorker(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(gvStageWorker.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			prevRow = gvStageWorker.GetSelectedRows()[0];
			StageWorkerModel model = (StageWorkerModel)StageWorkerBO.Instance.FindByPK(id);

			frmAddEditStageWorker frm = new frmAddEditStageWorker(cGlobalVariables.Edit);
			frm.stageWorkerModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadData();
				gvStageWorker.FocusedRowHandle = prevRow;
			}
		}

		private void btnDel_Click(object sender, EventArgs e)
		{
			if (!gvStageWorker.IsDataRow(gvStageWorker.FocusedRowHandle))
				return;
			int strID = TextUtils.ToInt(gvStageWorker.GetFocusedRowCellValue(colID));
			string str = TextUtils.ToString(gvStageWorker.GetFocusedRowCellValue("WorkerCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa công nhân [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
						StageWorkerBO.Instance.Delete(strID);
						gvStageWorker.DeleteRow(gvStageWorker.FocusedRowHandle);
						MessageBox.Show("Xoá thành công", TextUtils.Caption);
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
			LoadData();
		}

		private void gvStageWorker_DoubleClick(object sender, EventArgs e)
		{
			btnEdit_Click(null, null);
		}

		private void frmStageWorkerHyp_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		#endregion


	}
}
