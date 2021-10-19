using DevExpress.XtraCharts.GLGraphics;
using DevExpress.XtraGrid.Views.Grid;
using HP.Business;
using HP.Model;
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

	public partial class frmKnifeWorkingH : _Forms
	{
		//ProductWorkingModel ProductWorking = new ProductWorkingModel();
		ProductWorkingAuditModel ProductWorkingAudit = new ProductWorkingAuditModel();
		public int KnifeDetailListID = 0;
		public string ProductCode = "";
		public int Loai;

		public frmKnifeWorkingH()
		{

			InitializeComponent();
		}
		private void frmProductWorking_Load(object sender, EventArgs e)
		{
			this.Text += " - " + ProductCode;
			loadStep();
			if(Loai==1)
			{
				colFrequency.Visible = true;
			}	
			else
			{
				colFrequency.Visible = false;
			}

		}
		void loadStep()
		{
			if (Loai == 3)
			{
				DataTable dt = TextUtilsHP.Select(string.Format("SELECT * FROM dbo.KnifeStep WHERE ToolsID = {0} ORDER BY SortOrder", KnifeDetailListID));
				grdData.DataSource = dt;
			}
			else
			{
				DataTable dt = TextUtilsHP.Select(string.Format("SELECT * FROM dbo.KnifeStep WHERE KnifeDetailListID = {0} ORDER BY SortOrder", KnifeDetailListID));
				grdData.DataSource = dt;
			}

		}

		void loadWorking()
		{
			int stepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
			if (chkAll.Checked) stepID = 0;

			if (Loai == 3)
			{
				DataTable dt = new DataTable();

				//dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductWorking WHERE ProductStepID = {0} ORDER BY SortOrder", stepID));
				dt = TextUtilsHP.LoadDataFromSP("spGetProductWorkingTools_ByProductID", "A"
					, new string[] { "@KnifeDetailListID", "@ProductStepID" }
					, new object[] { this.KnifeDetailListID, stepID });

				grdDetail.DataSource = dt;
			}
			else
			{

				DataTable dt = new DataTable();

				//dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductWorking WHERE ProductStepID = {0} ORDER BY SortOrder", stepID));
				dt = TextUtilsHP.LoadDataFromSP("spGetProductWorking_ByProductID", "A"
					, new string[] { "@KnifeDetailListID", "@ProductStepID" }
					, new object[] { this.KnifeDetailListID, stepID });

				grdDetail.DataSource = dt;
			}

		}

		private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
		{
			loadWorking();
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			frmKnifeWorkingDetailH frm = new frmKnifeWorkingDetailH();
			frm.StepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
			frm.KnifeDetailListID = KnifeDetailListID;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadWorking();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
				if (id == 0) return;
				KnifeJigWorkingModel model = (KnifeJigWorkingModel)KnifeJigWorkingBO.Instance.FindByPK(id);
				frmKnifeWorkingDetailH frm = new frmKnifeWorkingDetailH();
				frm._KnifeJigWorking = model;
				frm.Loai = Loai;
				frm.KnifeDetailListID = this.KnifeDetailListID;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					loadWorking();
				}
			}
			catch (Exception ex)
			{
			}
		}

		private void grdDetail_DoubleClick(object sender, EventArgs e)
		{
			if (grvDetail.RowCount > 0)
				btnEdit_Click(null, null);
		}

		private void btnDelete_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
				if (id == 0) return;
				KnifeJigWorkingModel model = (KnifeJigWorkingModel)KnifeJigWorkingBO.Instance.FindByPK(id);
				if (MessageBox.Show("Bạn có chắc muốn xóa mục cần kiểm tra [" + grvDetail.GetFocusedRowCellValue(colWorkingName).ToString() + "] không?",
					  TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
					return;
				KnifeJigWorkingBO.Instance.Delete(id);
				grvDetail.DeleteSelectedRows();
				loadWorking();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void btnAddStep_Click(object sender, EventArgs e)
		{
			frmKnifeStepDetails frm = new frmKnifeStepDetails();
			//	frm._WorkingStepID = KnifeDetailListID;
			frm._KnifeDetailListID = KnifeDetailListID;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadStep();
			}
		}

		private void btnEditStep_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
				if (id == 0) return;
				KnifeStepModel model = (KnifeStepModel)KnifeStepBO.Instance.FindByPK(id);
				frmKnifeStepDetails frm = new frmKnifeStepDetails();
				frm._KnifeStep = model;
				frm._KnifeDetailListID = KnifeDetailListID;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					loadStep();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnDeleteStep_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
				if (id == 0) return;

				if (MessageBox.Show("Bạn có chắc muốn xóa công đoạn kiểm tra ["
					+ TextUtils.ToString(grvData.GetFocusedRowCellValue(colStepCode).ToString()) + " - "
					+ TextUtils.ToString(grvData.GetFocusedRowCellValue(colStepName).ToString()) + "] không?",
					  TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
					return;
				KnifeStepBO.Instance.Delete(id);
				//TextUtils.ExcuteSQL(string.Format("UPDATE dbo.ProductStep SET IsDeleted = 1 Where ID = {0}", id));
				grvData.DeleteSelectedRows();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0)
				btnEditStep_Click(null, null);
		}
		//pb053153

		private void btnSave_Click(object sender, EventArgs e)
		{
			grvDetail.FocusedRowHandle = -1;
			for (int i = 0; i < grvDetail.RowCount; i++)
			{
				int id = TextUtils.ToInt(grvDetail.GetRowCellValue(i, colWorkingID));
				if (id == 0) continue;
				string value = TextUtils.ToString(grvDetail.GetRowCellValue(i, colValue));
				decimal min = TextUtils.ToDecimal(grvDetail.GetRowCellValue(i, colMin));
				decimal max = TextUtils.ToDecimal(grvDetail.GetRowCellValue(i, colMax));
				string updateSql = string.Format(@"UPDATE dbo.KnifeJigWorking SET PeriodValue = '{0}' , MinValue = {1}, MaxValue = {2}                                     
                                        WHERE ID = {3}", value, min, max, id);
				TextUtilsHP.ExcuteSQL(updateSql);
			}
			//MessageBox.Show("Cất dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//loadWorking();
		}



		private void btnRestart_Click(object sender, EventArgs e)
		{
			loadWorking();
		}

		private void grvDetail_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
		{
			if (TextUtils.ToString(e.CellValue) == "Chờ duyệt")
			{
				e.Appearance.BackColor = Color.Yellow;
			}
		}

		private void btnRestartCD_Click(object sender, EventArgs e)
		{
			loadStep();
		}

		private void chkAll_CheckedChanged(object sender, EventArgs e)
		{
			loadWorking();
		}

		private void grdDetail_Click(object sender, EventArgs e)
		{

		}
	}
}
