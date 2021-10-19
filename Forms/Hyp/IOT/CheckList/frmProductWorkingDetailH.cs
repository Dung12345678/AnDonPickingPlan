using DevExpress.CodeParser;
using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.util;
using System.Windows.Forms;

namespace BMS
{

    public partial class frmProductWorkingDetailH : _Forms
    {
        public ProductWorkingModel ProductWorking = new ProductWorkingModel();
        public ProductWorkingModel ProductWorkingOld = new ProductWorkingModel();
        public ProductWorkingAuditModel ProductWorkingAudit = new ProductWorkingAuditModel();

        public int StepID = 0;
        public int ProductID = 0;
        public int _isApproved = 0;
        public int _actionType = 0;

        public frmProductWorkingDetailH()
        {
            InitializeComponent();
        }

        void LoadProductOld()
        {
            ProductWorkingOld.ID = ProductWorking.ID;
            ProductWorkingOld.ProductID = ProductWorking.ProductID;
            ProductWorkingOld.ProductStepID = ProductWorking.ProductStepID;
            ProductWorkingOld.WorkingName = ProductWorking.WorkingName;
            ProductWorkingOld.SortOrder = ProductWorking.SortOrder;
            ProductWorkingOld.Unit = ProductWorking.Unit;
            ProductWorkingOld.IsHidden = chkIsShow.Checked;
            ProductWorkingOld.ValueType = ProductWorking.ValueType;
            ProductWorkingOld.MaxValue = ProductWorking.MaxValue;
            ProductWorkingOld.MinValue = ProductWorking.MinValue;
            ProductWorkingOld.CheckValueType = ProductWorking.CheckValueType;
            ProductWorkingOld.ProductStepCode = ProductWorking.ProductStepCode;
            ProductWorkingOld.PeriodValue = ProductWorking.PeriodValue;
            ProductWorkingOld.CreatedBy = ProductWorking.CreatedBy;
        }
        private void frmProductWorkingDetail_Load(object sender, EventArgs e)
        {
            //ProductWorkingOld = ProductWorking;


            loadStep();
            loadData();
            LoadProductOld();
            if (ProductWorking.ID == 0)
            {
                cboStep.EditValue = StepID;
            }
            txtName.Focus();
        }
        void loadStep()
        {
            DataTable dt = new DataTable();
            dt = LibIE.Select(string.Format("SELECT * FROM dbo.ProductStep WHERE ProductID = {0} ORDER BY SortOrder", ProductID));
            cboStep.Properties.DataSource = dt;
            cboStep.Properties.DisplayMember = "ProductStepCode";
            cboStep.Properties.ValueMember = "ID";
        }

        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            if (ProductWorking.ID == 0)
            {
                txtSortOrder.Value = 1;
            }
            chkIsShow.Checked = ProductWorking.IsHidden;
            txtMax.Value = ProductWorking.MaxValue;
            txtMin.Value = ProductWorking.MinValue;
            txtName.Text = ProductWorking.WorkingName;
            txtPeriodValue.Text = ProductWorking.PeriodValue;
            txtSortOrder.Value = ProductWorking.SortOrder;
            txtUnit.Text = ProductWorking.Unit;
            cboValueType.SelectedIndex = ProductWorking.ValueType;
            cboStep.EditValue = ProductWorking.ProductStepID;
            cboCheckValueType.SelectedIndex = ProductWorking.CheckValueType;
        }

        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (cboStep.EditValue == null)
            {
                MessageBox.Show("Xin hãy chọn một công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (txtName.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền Mô tả chi tiết kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (cboValueType.SelectedIndex < 0)
            {
                MessageBox.Show("Xin hãy chọn một kiểu kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            string sql1 = "";
            if (ProductWorking.ID > 0)
            {
                sql1 = $"SELECT top 1 ID from ProductWorking WHERE SortOrder = '{txtSortOrder.Value}' and ID <>'{ProductWorking.ID}' and ProductID={ProductWorking.ProductID} and ProductStepID={ProductWorking.ProductStepID} ";
            }
            else
            {
                sql1 = $"SELECT top 1 ID from ProductWorking WHERE SortOrder = '{txtSortOrder.Value}' and ProductID={ProductID} and ProductStepID={StepID}";
            }
            int id = TextUtils.ToInt(LibIE.ExcuteScalar(sql1));
            if (id > 0)
            {
                MessageBox.Show("Trùng STT", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //return false;
            }
            if (cboCheckValueType.SelectedIndex <= 0)
            {
                MessageBox.Show("Xin hãy chọn một loại giá trị kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Cất dữ liệu
        /// </summary>
        /// <returns></returns>
        bool saveData()
        {
            if (!ValidateForm())
                return false;
            string reason = "";
            string sql = $"SELECT ReasonChange from ProductWorkingAudit " +
                                   $"WHERE ProductWorkingID='{ProductWorking.ID}' ";
            reason = TextUtils.ToString(LibIE.ExcuteScalar(sql));
            frmReasonChange frm = new frmReasonChange();
            frm.Reason = reason;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                reason = frm.Reason;
                if (ProductWorking.ID > 0)
                {
                    ProductWorking.ActionType = 2;
                    ProductWorking.IsApproved = 1;
                    //MessageBox.Show("Chờ duyệt", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ProductWorkingBO.Instance.Update(ProductWorking);
                }
                if (Global.UserID != 0)
                {
                    ProductWorking.CreatedBy = Global.AppFullName;
                    ProductWorking.UpdatedBy = Global.AppFullName;
                }
                ProductWorking.ProductID = ProductID;
                ProductWorking.ProductStepID = TextUtils.ToInt(cboStep.EditValue);
                ProductWorking.WorkingName = txtName.Text.Trim();
                ProductWorking.SortOrder = (int)txtSortOrder.Value;
                ProductWorking.IsHidden = chkIsShow.Checked;
                ProductWorking.Unit = txtUnit.Text.Trim();
                ProductWorking.ValueType = cboValueType.SelectedIndex;// kiểu kiểm tra
                ProductWorking.ValueTypeName = cboValueType.SelectedIndex == 0 ? "Check mark" : "Giá trị\n数値";
                ProductWorking.MaxValue = TextUtils.ToDecimal(txtMax.Value);
                ProductWorking.MinValue = TextUtils.ToDecimal(txtMin.Value);
                ProductWorking.CheckValueType = cboCheckValueType.SelectedIndex;
                ProductWorking.ProductStepCode = cboStep.Text;
                ProductWorking.IsApproved = 1;
                if (ProductWorking.ValueType == 0)//checkmark
                {
                    ProductWorking.PeriodValue = txtPeriodValue.Text.Trim();
                }
                else
                {
                    if (ProductWorking.CheckValueType == 1)//giá trị dạng số
                    {
                        if (ProductWorking.MinValue == ProductWorking.MaxValue)
                        {
                            ProductWorking.PeriodValue = ProductWorking.MaxValue.ToString("n3");
                        }
                        else
                        {
                            ProductWorking.PeriodValue = ProductWorking.MinValue.ToString("n3") + "~" + ProductWorking.MaxValue.ToString("n3");
                        }
                    }
                    else
                    {
                        ProductWorking.PeriodValue = txtPeriodValue.Text.Trim();
                    }
                }

                if (ProductWorking.ID > 0)
                {

                    if (_isApproved == 1 && _actionType == 1)
                    {
                        ProductWorking.ActionType = 1;
                        ProductWorkingBO.Instance.Update(ProductWorking);
                        return true;
                    }
                }
                else
                {
                    ProductWorking.ActionType = 1;
                    ProductWorking.ID = (int)ProductWorkingBO.Instance.Insert(ProductWorking);
                }
                processAudit(ProductWorking, reason);
                return true;
            }
            return false;
        }
        void processAudit(ProductWorkingModel ProductWorking, string reason)
        {
            //thêm or sửa vào bảng ProductWorkingAudit New
            ProductWorkingAudit.ProductWorkingID = ProductWorking.ID;
            ProductWorkingAudit.ProductID = ProductID;
            ProductWorkingAudit.ProductStepID = TextUtils.ToInt(cboStep.EditValue);
            ProductWorkingAudit.Unit = txtUnit.Text.Trim();
            ProductWorkingAudit.ReasonChange = reason;
            ProductWorkingAudit.IsApproved = 1;
            ProductWorkingAudit.ActionType = 1;

            if (ProductWorkingOld.ID > 0)
            {
                ProductWorkingAudit.ValueType = ProductWorkingOld.ValueType;// kiểu kiểm tra
                ProductWorkingAudit.MaxValue = ProductWorkingOld.MaxValue;
                ProductWorkingAudit.MinValue = ProductWorkingOld.MinValue;
                ProductWorkingAudit.WorkingName = ProductWorkingOld.WorkingName;
                ProductWorkingAudit.SortOrder = ProductWorkingOld.SortOrder;
                ProductWorkingAudit.CheckValueType = ProductWorkingOld.CheckValueType;
                ProductWorkingAudit.ProductStepCode = ProductWorkingOld.ProductStepCode;
                ProductWorkingAudit.PeriodValue = ProductWorkingOld.PeriodValue;
                ProductWorkingAudit.CreatedBy = ProductWorkingOld.CreatedBy;
                ProductWorkingAudit.IsHidden = ProductWorkingOld.IsHidden;
                ProductWorkingAudit.ActionType = 2;

            }
            ProductWorkingAudit.ValueTypeNew = ProductWorking.ValueType;// kiểu kiểm tra
            ProductWorkingAudit.UpdatedBy = ProductWorking.UpdatedBy;
            ProductWorkingAudit.IsHiddenNew = ProductWorking.IsHidden;
            ProductWorkingAudit.MaxValueNew = ProductWorking.MaxValue;
            ProductWorkingAudit.MinValueNew = ProductWorking.MinValue;
            ProductWorkingAudit.WorkingNameNew = ProductWorking.WorkingName;
            ProductWorkingAudit.SortOrderNew = ProductWorking.SortOrder;
            ProductWorkingAudit.CheckValueTypeNew = ProductWorking.CheckValueType;
            ProductWorkingAudit.ProductStepCodeNew = ProductWorking.ProductStepCode;
            ProductWorkingAudit.PeriodValueNew = ProductWorking.PeriodValue;
            ProductWorkingAudit.UpdatedBy = ProductWorking.UpdatedBy;
            ProductWorkingAuditBO.Instance.Insert(ProductWorkingAudit);
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ProductWorkingOld.ID > 0 &&
             ProductWorkingOld.ID == ProductWorking.ID &&
            ProductWorkingOld.ProductID == ProductID &&
            ProductWorkingOld.ProductStepID == TextUtils.ToInt(cboStep.EditValue) &&
            ProductWorkingOld.WorkingName == txtName.Text.Trim() &&
            ProductWorkingOld.SortOrder == (int)txtSortOrder.Value &&
            ProductWorkingOld.Unit == txtUnit.Text.Trim() &&
            ProductWorkingOld.IsHidden == chkIsShow.Checked &&
            ProductWorkingOld.ValueType == cboValueType.SelectedIndex &&
            ProductWorkingOld.MaxValue == TextUtils.ToDecimal(txtMax.Value) &&
            ProductWorkingOld.MinValue == TextUtils.ToDecimal(txtMin.Value) &&
            ProductWorkingOld.CheckValueType == cboCheckValueType.SelectedIndex &&
            ProductWorkingOld.ProductStepCode == cboStep.Text.Trim() &&
            ProductWorkingOld.PeriodValue == txtPeriodValue.Text.Trim())
            {
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                if (saveData())
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            if (saveData())
            {
                ProductWorking = new ProductWorkingModel();
                loadData();
                txtName.Focus();
            }
        }
        private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
        private void txtMin_ValueChanged(object sender, EventArgs e)
        {
            txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
        }
        private void txtMax_ValueChanged(object sender, EventArgs e)
        {
            txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSave_Click(null, null);
        }

        private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSaveNew_Click(null, null);
        }

        private void cboStep_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void cboValueType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
