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
    public partial class frmWorkingH : _Forms
    {
        public frmWorkingH()
        {
            InitializeComponent();
        }

        private void frmWorking_Load(object sender, EventArgs e)
        {
            loadGroup();
        }

        void loadGroup()
        {
            DataTable dt = LibIE.Select("SELECT * FROM dbo.ProductGroup WITH(NOLOCK)");
            cboGroup.Properties.DataSource = dt;
            cboGroup.Properties.DisplayMember = "ProductGroupName";
            cboGroup.Properties.ValueMember = "ID";
        }

        void loadStep()
        {
            int groupID = TextUtils.ToInt(cboGroup.EditValue);
            DataTable dt = new DataTable();
            if (groupID > 0)
            {
                dt = LibIE.Select(string.Format("SELECT * FROM dbo.WorkingStep WHERE ProductGroupID = {0} AND (IsDeleted = 0 OR IsDeleted IS NULL) ORDER BY SortOrder", groupID));
            }
            grdData.DataSource = dt;
        }

        void loadWorking()
        {
            int stepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            DataTable dt = new DataTable();
            if (chkAll.Checked)
            {
                dt = LibIE.Select(string.Format($"select w.*,(CAST(ws.SortOrder as nvarchar)+'-'+cast(ws.WorkingStepCode as nvarchar)) as WorkingStepCodes, case when w.CheckValueType=1 then N'Số' when w.CheckValueType = 2 then N'Ký tự' when w.CheckValueType = 0 then ''  END as CheckValueTypes from Working w left join WorkingStep ws on w.WorkingStepID = ws.ID Left join ProductGroup pg on ws.ProductGroupID = pg.ID where ProductGroupID = '{cboGroup.EditValue}' Order by w.SortOrder"));
            }
            else
            {
                if (stepID > 0)
                {
                    dt = LibIE.Select(string.Format("SELECT w.* ,case when w.CheckValueType=1 then N'Số' when w.CheckValueType = 2 then N'Ký tự' when w.CheckValueType = 0 then '' END as CheckValueTypes FROM dbo.Working w WHERE WorkingStepID = {0} ORDER BY SortOrder", stepID));
                }
            }
           
            grdDetail.DataSource = dt;
        }

        private void cboGroup_EditValueChanged(object sender, EventArgs e)
        {
            loadStep();
        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadWorking();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmWorkingDetailH frm = new frmWorkingDetailH();
            frm.GroupID = TextUtils.ToInt(cboGroup.EditValue);
            frm.StepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
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
                WorkingModel model = (WorkingModel)WorkingBO.Instance.GetObjectByID(id);
                frmWorkingDetailH frm = new frmWorkingDetailH();
                frm.WorkingModel = model;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    loadWorking();
                }
            }
            catch (Exception)
            {
            }
        }
        private void grdDetail_DoubleClick(object sender, EventArgs e)
        {
            if (grvDetail.RowCount > 0 && btnEdit.Enabled == true)
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
                WorkingBO.Instance.Delete(id);
                grvDetail.DeleteSelectedRows();
                //Xóa các thằng con thuộc thằng cha này đi
                LibIE.ExcuteSQL(string.Format("EXEC spUpdateMultiProductWorking {0}, 2", id));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            frmHistory frm = new frmHistory();
            frm.ShowDialog();
        }

        private void Addcopy_Click(object sender, EventArgs e)
        {
            frmCopyWorkingToGroupH frm = new frmCopyWorkingToGroupH();
            frm.GroupID = TextUtils.ToInt(cboGroup.EditValue);
            frm.WorkingStepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            frm.WorkingID = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
            frm.GroupName = cboGroup.Text.Trim();
            frm.WorkingStepCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colStepCode));
            frm.STT= TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colSortOrder));
            frm.WorkingName= TextUtils.ToString(grvDetail.GetFocusedRowCellValue(colWorkingName));
            if (TextUtils.ToInt(cboGroup.EditValue) == 0 || TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID)) == 0 || TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID)) == 0) return;
                frm.ShowDialog();
        }

        private void btnShowWorkingStep_Click(object sender, EventArgs e)
        {
            frmCopyStepH frm = new frmCopyStepH();
            frm.GroupID = TextUtils.ToInt(cboGroup.EditValue);
            frm.WorkingStepID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID));
            frm.GroupName = cboGroup.Text.Trim();
            frm.StepCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colStepCode));
            frm.StepName = TextUtils.ToString(grvData.GetFocusedRowCellValue(colStepName));
            frm.SortOrder = TextUtils.ToString(grvData.GetFocusedRowCellValue(colSortOrderStep));
            if (TextUtils.ToInt(cboGroup.EditValue)==0 || TextUtils.ToInt(grvData.GetFocusedRowCellValue(colStepID))==0) return;
            frm.ShowDialog();
        }

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            loadWorking();
        }
    }
}
