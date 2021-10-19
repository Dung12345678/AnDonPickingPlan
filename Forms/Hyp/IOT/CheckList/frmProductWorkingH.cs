using DevExpress.XtraCharts.GLGraphics;
using DevExpress.XtraGrid.Views.Grid;
using IE.Business;
using IE.Model;
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

    public partial class frmProductWorkingH : _Forms
    {
        //ProductWorkingModel ProductWorking = new ProductWorkingModel();
        ProductWorkingAuditModel ProductWorkingAudit = new ProductWorkingAuditModel();
        public int ProductID = 0;
        public string ProductCode = "";


        public frmProductWorkingH()
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
            DataTable dt = LibIE.Select(string.Format("SELECT * FROM dbo.ProductStep WHERE ProductID = {0} AND IsDeleted = 0 ORDER BY SortOrder", ProductID));
            grdData.DataSource = dt;
        }

        void loadWorking()
        {
            int stepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            if (chkAll.Checked) stepID = 0;
            DataTable dt = new DataTable();

            //dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductWorking WHERE ProductStepID = {0} ORDER BY SortOrder", stepID));
            dt = LibIE.LoadDataFromSP("spGetProductWorking_ByProductID", "A"
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
            frmProductWorkingDetailH frm = new frmProductWorkingDetailH();
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
                int isApproved = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colIsApproved));
                int actionType = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colActionType));
                string actionTypes = TextUtils.ToString(grvDetail.GetFocusedRowCellValue(colActionTypes));
                if (isApproved == 1 && (actionType == 3 || actionType == 2))
                {
                    MessageBox.Show("Mục kiểm tra đang chờ duyệt", "Không được sửa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
                if (id == 0) return;
                ProductWorkingModel model = (ProductWorkingModel)ProductWorkingBO.Instance.FindByPK(id);
                frmProductWorkingDetailH frm = new frmProductWorkingDetailH();
                frm.ProductWorking = model;
                frm.ProductID = this.ProductID;
                frm._isApproved = isApproved;
                frm._actionType = actionType;
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
                int isApprored = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colIsApproved));
                int actionType = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colActionType));
                string ActionType = TextUtils.ToString(grvDetail.GetFocusedRowCellValue(colActionTypes));
                int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
                //      int ProductWorkingID = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue());
                if (id == 0) return;
                if ((actionType == 2 || actionType == 3) && isApprored == 1)
                {
                    MessageBox.Show("Mục kiểm tra đang chờ duyệt", "Không được xóa", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ProductWorkingModel model = (ProductWorkingModel)ProductWorkingBO.Instance.FindByPK(id);
                if (MessageBox.Show("Bạn có chắc muốn xóa mục cần kiểm tra [" + grvDetail.GetFocusedRowCellValue(colWorkingName).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                // Chờ duyệt và Thêm thì xóa trực tiếp không ghi vào Lịch sử hành động
                if (TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colIsApproved)) == 1 && TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colActionType)) == 1)
                {
                    ProductWorkingAuditBO.Instance.DeleteByAttribute("ProductWorkingID", id);
                    ProductWorkingBO.Instance.Delete(id);
                    grvDetail.DeleteSelectedRows();
                }
                else
                {
                    // MessageBox.Show("Hành động này cần được duyệt", "Xác nhận", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (model.ID > 0)
                    {
                        model.IsApproved = 1;
                        model.ActionType = 3;
                        ProductWorkingBO.Instance.Update(model);
                        string reason = "";
                        frmReasonChange frm = new frmReasonChange();
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            reason = frm.Reason;
                        }
                        ProductWorkingAudit.ProductID = ProductID;
                        ProductWorkingAudit.ProductWorkingID = model.ID;
                        ProductWorkingAudit.ProductStepCode = model.ProductStepCode;
                        ProductWorkingAudit.MaxValue = model.MaxValue;
                        ProductWorkingAudit.MinValue = model.MinValue;
                        ProductWorkingAudit.WorkingName = model.WorkingName;
                        ProductWorkingAudit.PeriodValue = model.PeriodValue;
                        ProductWorkingAudit.SortOrder = model.SortOrder;
                        ProductWorkingAudit.Unit = model.Unit;
                        ProductWorkingAudit.ValueType = model.ValueType;
                        ProductWorkingAudit.ProductStepID = model.ProductStepID;
                        ProductWorkingAudit.CheckValueType = model.CheckValueType;
                        ProductWorkingAudit.CreatedBy = Global.AppFullName;
                        ProductWorkingAudit.CreatedDate = DateTime.Now;
                        ProductWorkingAudit.UpdatedBy = Global.AppFullName;
                        ProductWorkingAudit.UpdatedDate = DateTime.Now;
                        ProductWorkingAudit.IsApproved = 1;
                        ProductWorkingAudit.ActionType = 3;
                        ProductWorkingAudit.ReasonChange = reason;
                        ProductWorkingAuditBO.Instance.Insert(ProductWorkingAudit);
                    }
                }
                loadWorking();
                //    ProductWorkingBO.Instance.Delete(id);
                //   grvDetail.DeleteSelectedRows();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btnAddStep_Click(object sender, EventArgs e)
        {
            frmProductStepDetailH frm = new frmProductStepDetailH();
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
                frmProductStepDetailH frm = new frmProductStepDetailH();
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
                LibIE.ExcuteSQL(updateSql);
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
    }
}
