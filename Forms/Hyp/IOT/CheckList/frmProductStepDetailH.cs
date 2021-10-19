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
    public partial class frmProductStepDetailH : _Forms
    {
        public ProductStepModel ProductStep = new ProductStepModel();

        public int ProductID = 0;

        public frmProductStepDetailH()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            txtSortOrder.Value = ProductStep.SortOrder;
            txtCode.Text = ProductStep.ProductStepCode;
            txtDes.Text = ProductStep.Description;
        }

        private void frmProductStepDetail_Load(object sender, EventArgs e)
        {
            loadData();
        }

        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (txtSortOrder.Value <= 0 )
            {
                MessageBox.Show("Xin hãy điền Thứ tự của công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (txtCode.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền Mã cho công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (txtDes.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền Mô tả cho công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
            ProductStep.ProductID = ProductID;
            ProductStep.ProductStepCode = txtCode.Text.Trim();
            ProductStep.SortOrder = (int)txtSortOrder.Value;
            ProductStep.Description = txtDes.Text.Trim();
            ProductStep.UpdatedBy = Global.AppUserName;
            ProductStep.UpdatedDate = DateTime.Now;

            if (ProductStep.ID <= 0)
            {
                ProductStep.CreatedBy = Global.AppUserName;
                ProductStep.CreatedDate = DateTime.Now;

                ProductStepBO.Instance.Insert(ProductStep);
            }
            else
            {
                ProductStepBO.Instance.Update(ProductStep);
            }
            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if(saveData())
            this.DialogResult = DialogResult.OK;
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            if (saveData())
            {
                ProductStep = new ProductStepModel();
                loadData();
            }
        }

        private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSave_Click(null, null);
        }

        private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
