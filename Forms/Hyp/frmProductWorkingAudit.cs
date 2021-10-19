using DevExpress.XtraGrid.Views.Grid;
using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmProductWorkingAudit : _Forms
    {

        public frmProductWorkingAudit()
        {
            InitializeComponent();
        }

        private void frmCheckworking_Load(object sender, EventArgs e)
        {
            dtpFrom.Value = DateTime.Now;
            dtpTo.Value = DateTime.Now;
            LoadInfoSearch();
            loadcboApprovedBy();
            //  cboIsApproveds.Properties.DisplayMember = "Chờ duyệt";
        }
        void LoadInfoSearch()
        {
            dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
            dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

            DataTable dt = LibIE.LoadDataFromSP("spGetCheckWorkingFindDate"
                        , "A"
                        , new string[] { "@DateStart", "@DateEnd ", "@TextFilter", "@CboIsApproved", "@UserApproved" }
                        , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
                                        , dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
                                        , txtAssyOrderId.Text.Trim()
                                        ,cboIsApproveds.SelectedIndex + 1
                                        ,cboUserApproved.Text}
                    );
            grdData.DataSource = dt;
        }

        private void btnFindDate_Click(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }
        void loadcboApprovedBy()
        {
            DataTable dt = TextUtils.Select("SELECT * FROM dbo.Users where ApprovedBy = 1");
            cboUserApproved.Properties.DataSource = dt;
            cboUserApproved.Properties.DisplayMember = "FullName";
            cboUserApproved.Properties.ValueMember = "ID";

            //int groupID = TextUtils.ToInt(cboApprovedBy.EditValue);
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            int isApproved = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colIsApproved));
            if (isApproved == 2) return;
            if (cboUserApproved.Text == "")
            {
                MessageBox.Show("Người duyệt không được để trống", "Xác nhận", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogResult Approve = MessageBox.Show("Bạn có chắc chắn duyệt", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Approve == DialogResult.No) return;
            Int32[] selectedRowHandles = grvData.GetSelectedRows();
            for (int i = 0; i < selectedRowHandles.Length; i++)
            {
                int selectedRowHandle = selectedRowHandles[i];
                if (selectedRowHandle >= 0)
                {
                    int id = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colID));
                    if (id == 0) return;
                    int ProductWorkingID = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colProductWorkingID));
                    if (ProductWorkingID == 0) return;
                    ProductWorkingModel productWorking = (ProductWorkingModel)ProductWorkingBO.Instance.FindByPK(ProductWorkingID);
                    ProductWorkingModel productWorkings = new ProductWorkingModel();
                    ProductWorkingAuditModel productWorkingAudit = (ProductWorkingAuditModel)ProductWorkingAuditBO.Instance.FindByPK(id);
                    isApproved = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colIsApproved));
                    if (isApproved == 2) return;
                    productWorkingAudit.IsApproved = 2;
                    productWorkingAudit.UserApproved = cboUserApproved.Text.Trim();
                    ProductWorkingAuditBO.Instance.Update(productWorkingAudit);
                    if (productWorkingAudit.ActionType == 1)
                    {
                        if (isApproved == 1)
                        {
                            productWorking.IsApproved = 2;
                            ProductWorkingBO.Instance.Update(productWorking);
                        }
                        if (isApproved == 3)
                        {
                            productWorkingAudit.IsApproved = 2;
                            productWorkingAudit.UserApproved = cboUserApproved.Text.Trim();
                            ProductWorkingAuditBO.Instance.Update(productWorkingAudit);
                            productWorkings.ProductID = productWorkingAudit.ProductID;
                            productWorkings.ProductStepID = productWorkingAudit.ProductStepID;
                            productWorkings.WorkingID = productWorkingAudit.WorkingID;
                            productWorkings.WorkingName = productWorkingAudit.WorkingNameNew;
                            productWorkings.SortOrder = productWorkingAudit.SortOrderNew;
                            productWorkings.Unit = productWorkingAudit.Unit;
                            productWorkings.ValueType = productWorkingAudit.ValueTypeNew;
                            productWorkings.ValueTypeName = TextUtils.ToString(grvData.GetRowCellValue(selectedRowHandle, colValueTypeNews));
                            productWorkings.IsHidden = TextUtils.ToBoolean(grvData.GetRowCellValue(selectedRowHandle, colIsHiddenNew));
                            productWorkings.MaxValue = productWorkingAudit.MaxValueNew;
                            productWorkings.MinValue = productWorkingAudit.MinValueNew;
                            productWorkings.CheckValueType = productWorkingAudit.CheckValueTypeNew;
                            productWorkings.ProductStepCode = productWorkingAudit.ProductStepCodeNew;
                            productWorkings.PeriodValue = productWorkingAudit.PeriodValueNew;
                            productWorkings.CreatedBy = productWorkingAudit.CreatedBy;
                            productWorkings.UpdatedBy = productWorkingAudit.UpdatedBy;
                            productWorkings.CreatedDate = productWorkingAudit.CreatedDate;
                            productWorkings.UpdatedDate = productWorkingAudit.UpdatedDate;
                            productWorkings.IsApproved = 2;
                            productWorkings.ActionType = 1;
                            ProductWorkingBO.Instance.Insert(productWorkings);
                        }
                    }
                    else if (productWorkingAudit.ActionType == 2)
                    {
                        productWorking.ProductID = productWorkingAudit.ProductID;
                        productWorking.ProductStepID = productWorkingAudit.ProductStepID;
                        productWorking.WorkingID = productWorkingAudit.WorkingID;
                        productWorking.WorkingName = productWorkingAudit.WorkingNameNew;
                        productWorking.SortOrder = productWorkingAudit.SortOrderNew;
                        productWorking.Unit = productWorkingAudit.Unit;
                        productWorking.IsHidden = productWorkingAudit.IsHiddenNew;
                        productWorking.ValueType = productWorkingAudit.ValueTypeNew;
                        productWorking.MaxValue = productWorkingAudit.MaxValueNew;
                        productWorking.MinValue = productWorkingAudit.MinValueNew;
                        productWorking.CheckValueType = productWorkingAudit.CheckValueTypeNew;
                        productWorking.ProductStepCode = productWorkingAudit.ProductStepCodeNew;
                        productWorking.PeriodValue = productWorkingAudit.PeriodValueNew;
                        productWorking.CreatedBy = productWorkingAudit.CreatedBy;
                        productWorking.UpdatedBy = productWorkingAudit.UpdatedBy;
                        productWorking.CreatedDate = productWorkingAudit.CreatedDate;
                        productWorking.UpdatedDate = productWorkingAudit.UpdatedDate;
                        productWorking.IsApproved = 2;
                        productWorking.ActionType = 2;
                        ProductWorkingBO.Instance.Update(productWorking);
                    }
                    else if (productWorkingAudit.ActionType == 3)
                    {
                        ProductWorkingBO.Instance.Delete(ProductWorkingID);
                    }
                }
            }
            LoadInfoSearch();
        }
        private void btnNoBrowser_Click(object sender, EventArgs e)
        { 
            int isApproved = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colIsApproved));
            if (isApproved == 2 || isApproved == 3) return;
            if (cboUserApproved.Text == "")
            {
                MessageBox.Show("Người duyệt không được để trống", "Xác nhận", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogResult NoApprove = MessageBox.Show("Bạn chắc chắn không duyệt", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (NoApprove == DialogResult.No) return;
            Int32[] selectedRowHandles = grvData.GetSelectedRows();
            for (int i = 0; i < selectedRowHandles.Length; i++)
            {

                int selectedRowHandle = selectedRowHandles[i];
                if (selectedRowHandle >= 0)
                {
                    isApproved = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colIsApproved));
                    if (isApproved == 2 || isApproved == 3) return;
                    int id = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colID));
                    if (id == 0) return;
                    ProductWorkingAuditModel model = (ProductWorkingAuditModel)ProductWorkingAuditBO.Instance.FindByPK(id);
                    model.IsApproved = 3;
                    model.UserApproved = cboUserApproved.Text.Trim();
                    ProductWorkingAuditBO.Instance.Update(model);
                    int ProductWorkingID = TextUtils.ToInt(grvData.GetRowCellValue(selectedRowHandle, colProductWorkingID));
                    if (ProductWorkingID == 0) return;
                    ProductWorkingModel productWorking = (ProductWorkingModel)ProductWorkingBO.Instance.FindByPK(ProductWorkingID);
                    if (model.ActionType == 1)
                    {
                        ProductWorkingBO.Instance.Delete(ProductWorkingID);
                    }
                    else if (model.ActionType == 2)
                    {
                        productWorking.IsApproved = 2;
                        ProductWorkingBO.Instance.Update(productWorking);
                    }
                    else if (model.ActionType == 3)
                    {
                        productWorking.IsApproved = 3;
                        ProductWorkingBO.Instance.Update(productWorking);
                    }
                }
            }
            LoadInfoSearch();
        }
        private void grvData_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string actionType = Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colAction));
            if (actionType == "Sửa")//Điều kiện
            {
                if (e.Column == colPeriodValue || e.Column == colPeriodValueNew)
                {
                    if (Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colPeriodValue)) != Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colPeriodValueNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colIsApproveds && TextUtils.ToInt(grvData.GetRowCellValue(e.RowHandle, colIsApproveds)) == 1)
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                if (e.Column == colSortOrder || e.Column == colSortOrderNew)
                {
                    if (Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colSortOrderNew)) != Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colSortOrder)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colProductStepCode || e.Column == colProductStepCodeNew)
                {
                    if (Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colProductStepCode)) != Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colProductStepCodeNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colWorkingName || e.Column == colWorkingNameNew)
                {
                    if (Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colWorkingName)) != Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colWorkingNameNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colValueTypeName || e.Column == colValueTypeNews)
                {
                    if (Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colValueTypeName)) != Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colValueTypeNews)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colCheckValueTypes || e.Column == colCheckValueTypeNews)
                {
                    if (Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colCheckValueTypes)) != Lib.ToString(grvData.GetRowCellValue(e.RowHandle, colCheckValueTypeNews)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colMinValue || e.Column == colMinValueNew)
                {
                    if (Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colMinValue)) != Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colMinValueNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colMaxValue || e.Column == colMaxValueNew)
                {
                    if (Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colMaxValue)) != Lib.ToInt(grvData.GetRowCellValue(e.RowHandle, colMaxValueNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                if (e.Column == colIsHidden || e.Column == colIsHiddenNew)
                {
                    if (Lib.ToBoolean(grvData.GetRowCellValue(e.RowHandle, colIsHidden)) != Lib.ToBoolean(grvData.GetRowCellValue(e.RowHandle, colIsHiddenNew)))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
            }
        }

		private void txtAssyOrderId_KeyDown(object sender, KeyEventArgs e)
		{
            if (e.KeyCode != Keys.Enter) return;
            LoadInfoSearch();
        }
	}
}

