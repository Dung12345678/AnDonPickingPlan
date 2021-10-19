using BMS.Business;
using BMS.Model;
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
    public partial class frmProductWorking : _Forms
    {
        public int ProductID = 0;

        public string ProductCode = "";
        public frmProductWorking()
        {
            InitializeComponent();
        }

        private void frmProductWorking_Load(object sender, EventArgs e)
        {
            this.Text += " - " + ProductCode;
            loadStep();
        }

        void loadStep()
        {
            DataTable dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductStep WHERE ProductID = {0} AND IsDeleted = 0 ORDER BY SortOrder", ProductID));
            grdData.DataSource = dt;
        }

        void loadWorking()
        {
            int stepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            if (chkAll.Checked) stepID = 0;
            DataTable dt = new DataTable();

            //dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductWorking WHERE ProductStepID = {0} ORDER BY SortOrder", stepID));
            dt = TextUtils.LoadDataFromSP("spGetProductWorking_ByProductID", "A"
                , new string[] { "@ProductID", "@ProductStepID" }
                , new object[] { this.ProductID, stepID });

            grdDetail.DataSource = dt;
        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadWorking();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmProductWorkingDetail frm = new frmProductWorkingDetail();
            frm.StepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            frm.ProductID = ProductID;
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
                ProductWorkingModel model = (ProductWorkingModel)ProductWorkingBO.Instance.FindByPK(id);
                frmProductWorkingDetail frm = new frmProductWorkingDetail();
                frm.ProductWorking = model;
                frm.ProductID = this.ProductID;
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

                if (MessageBox.Show("Bạn có chắc muốn xóa mục cần kiểm tra [" + grvDetail.GetFocusedRowCellValue(colWorkingName).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                ProductWorkingBO.Instance.Delete(id);
                grvDetail.DeleteSelectedRows();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddStep_Click(object sender, EventArgs e)
        {
            frmProductStepDetail frm = new frmProductStepDetail();
            frm.ProductID = ProductID;
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
                ProductStepModel model = (ProductStepModel)ProductStepBO.Instance.FindByPK(id);
                frmProductStepDetail frm = new frmProductStepDetail();
                frm.ProductStep = model;
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
                ProductStepBO.Instance.Delete(id);
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

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            loadWorking();
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
                string updateSql = string.Format(@"UPDATE dbo.ProductWorking SET PeriodValue = '{0}' , MinValue = {1}, MaxValue = {2}                                     
                                        WHERE ID = {3}", value, min, max, id);
                TextUtils.ExcuteSQL(updateSql);
            }
            //MessageBox.Show("Cất dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //loadWorking();
        }

        private void btnRestartCD_Click(object sender, EventArgs e)
        {
            loadStep();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            loadWorking();
        }
    }
}
