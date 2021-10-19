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
    public partial class frmProductWorkingDetail : _Forms
    {
        public ProductWorkingModel ProductWorking = new ProductWorkingModel();
        public int StepID = 0;
        public int ProductID = 0;

        public frmProductWorkingDetail()
        {
            InitializeComponent();
        }
        private void frmProductWorkingDetail_Load(object sender, EventArgs e)
        {

            loadStep();
            loadData();
            if (ProductWorking.ID == 0)
            {
                cboStep.EditValue = StepID;
            }
            txtName.Focus();
        }

        void loadStep()
        {
            DataTable dt = new DataTable();
            dt = TextUtils.Select(string.Format("SELECT * FROM dbo.ProductStep WHERE ProductID = {0} ORDER BY SortOrder", ProductID));
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
            txtMax.Value = ProductWorking.MaxValue;
            txtMin.Value = ProductWorking.MinValue;
            txtName.Text = ProductWorking.WorkingName;
            txtPeriodValue.Text = ProductWorking.PeriodValue;
            txtSortOrder.Value = ProductWorking.SortOrder;
            txtUnit.Text = ProductWorking.Unit;
            chkIsHidden.Checked = ProductWorking.IsHidden;
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
            string sql = "";
            if (ProductWorking.ID > 0)
            {
                sql = $"select ID from ProductWorking pw where pw.SortOrder={txtSortOrder.Value} and pw.ID <> {ProductWorking.ID} and pw.ProductStepID={TextUtils.ToInt(cboStep.EditValue)} and pw.ProductID={ProductID}";
            }
            else
            {
                sql = $"select ID from ProductWorking pw where pw.SortOrder={txtSortOrder.Value} and pw.ProductStepID={TextUtils.ToInt(cboStep.EditValue)} and pw.ProductID={ProductID}";
            }
            int id = TextUtils.ToInt(TextUtils.ExcuteScalar(sql));
            if (id > 0)
            {
                MessageBox.Show("Trùng STT", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                // return false;
            }

            //if (cboCheckValueType.SelectedIndex <= 0)
            //{
            //    MessageBox.Show("Xin hãy chọn một loại giá trị kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return false;
            //}
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

            ProductWorking.ProductID = ProductID;
            ProductWorking.ProductStepID = TextUtils.ToInt(cboStep.EditValue);
            ProductWorking.WorkingName = txtName.Text.Trim();
            ProductWorking.SortOrder = (int)txtSortOrder.Value;
            ProductWorking.Unit = txtUnit.Text.Trim();
            ProductWorking.ValueType = cboValueType.SelectedIndex;
            ProductWorking.ValueTypeName = cboValueType.SelectedIndex == 0 ? "Check mark" : "Giá trị\n数値";
            ProductWorking.MaxValue = TextUtils.ToDecimal(txtMax.Value);
            ProductWorking.MinValue = TextUtils.ToDecimal(txtMin.Value);
            ProductWorking.IsHidden = chkIsHidden.Checked;
            ProductWorking.CheckValueType = cboCheckValueType.SelectedIndex;
            ProductWorking.ProductStepCode = cboStep.Text;
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
                ProductWorkingBO.Instance.Update(ProductWorking);
            }
            else
            {
                ProductWorking.ID = (int)ProductWorkingBO.Instance.Insert(ProductWorking);
            }
            return true;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (saveData())
                this.DialogResult = DialogResult.OK;
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
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
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
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
