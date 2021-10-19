using BMS.Business;
using BMS.Model;
using BMS.Utils;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmCheckForceDetail : _Forms
    {
        public ProductCheckForceModel model = new ProductCheckForceModel();

        public frmCheckForceDetail()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            txtProductID.Text = model.ProductCode;
            setTextBaseProductCode(txtProductID.Text);
            nmrCheckForceMin.Value = model.CheckForceMin;
            nmrForceCheckMax.Value = model.CheckForceMax;
            nmrGunNum.Value = model.GunNumber;
            nmrJob.Value = model.JobNumber;
            nmrQtyBulong.Value = model.QtyBuLong;
            nmrShootNum.Value = model.ShootNumber;
            nmrBanGa.Value = model.QtyOcBanGa;
            nmrBanThat.Value = model.QtyOcBanThat;
            nmrSetForce.Value = model.SetForce;
            if (string.IsNullOrEmpty(model.ProductStepCode))
                cbbProductStepCode.SelectedIndex = 0;
            else
                cbbProductStepCode.SelectedText = model.ProductStepCode;
        }

        private void frmProductStepDetail_Load(object sender, EventArgs e)
        {
            cbbProductStepCode.Items.Add("-Chọn công đoạn-");
            loadData();
        }


        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            string productCode = txtProductID.Text.Trim();
            if (string.IsNullOrEmpty(productCode))
            {
                MessageBox.Show("Bạn chưa nhập mã sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            Expression expProduct = new Expression("ProductCode", productCode);
            ArrayList arrProduct = ProductBO.Instance.FindByExpression(expProduct);
            if (arrProduct.Count <= 0)
            {
                MessageBox.Show("Mã sản phẩm không tồn tại", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (cbbProductStepCode.SelectedIndex == 0)
            {
                MessageBox.Show("Bạn chưa chọn công đoạn", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            model.ProductCode = txtProductID.Text;
            model.CheckForceMin = nmrCheckForceMin.Value;
            model.CheckForceMax = nmrForceCheckMax.Value;
            model.GunNumber = TextUtils.ToInt(nmrGunNum.Value);
            model.JobNumber = TextUtils.ToInt(nmrJob.Value);
            model.QtyBuLong = TextUtils.ToInt(nmrQtyBulong.Value);
            model.ShootNumber = TextUtils.ToInt(nmrShootNum.Value);
            model.QtyOcBanGa = TextUtils.ToInt(nmrBanGa.Value);
            model.QtyOcBanThat = TextUtils.ToInt(nmrBanThat.Value);
            model.SetForce = nmrSetForce.Value;
            model.ProductStepCode = cbbProductStepCode.SelectedText;
            if (model.ID > 0)
            {
                ProductCheckForceBO.Instance.Update(model);
            }
            else
            {
                ProductCheckForceBO.Instance.Insert(model);
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
            if (saveData()) { 
                model = new ProductCheckForceModel();
                loadData();
            }
        }
        private void setTextBaseProductCode(string ProductCode)
        {
            DataTable dt = ProductBO.Instance.LoadDataFromSP("spGetProductStep_ByProductCode", "Table"
                    , new string[] { "@ProductCode" }, new object[] { ProductCode });
            if (dt.Rows.Count <= 0)
            {
                txtProductName.Text = string.Empty;
                txtProductGroupCode.Text = string.Empty;
                return;
            }
            txtProductName.Text = TextUtils.ToString(dt.Rows[0]["ProductName"]);
            txtProductGroupCode.Text = TextUtils.ToString(dt.Rows[0]["ProductGroupCode"]);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (!cbbProductStepCode.Items.Contains(dt.Rows[i]["ProductStepCode"]))
                    cbbProductStepCode.Items.Add(dt.Rows[i]["ProductStepCode"]);
            }
        }

        private void txtProductID_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                setTextBaseProductCode(txtProductID.Text);
            }
        }

        private void nmrShootNum_ValueChanged(object sender, EventArgs e)
        {
            if (nmrShootNum.Value - nmrBanGa.Value > 0)
                nmrBanThat.Value = nmrShootNum.Value - nmrBanGa.Value;
            else
                nmrBanThat.Value = 0;
        }

        private void nmrBanGa_ValueChanged(object sender, EventArgs e)
        {
            if (nmrShootNum.Value - nmrBanGa.Value > 0)
                nmrBanThat.Value = nmrShootNum.Value - nmrBanGa.Value;
            else
                nmrBanThat.Value = 0;
        }
    }
}
