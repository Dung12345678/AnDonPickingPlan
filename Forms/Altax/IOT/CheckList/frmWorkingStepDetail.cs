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
    public partial class frmWorkingStepDetail : _Forms
    {
        public WorkingStepModel WorkingStep = new WorkingStepModel();
        public frmWorkingStepDetail()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            txtSortOrder.Value = WorkingStep.SortOrder;
            txtCode.Text = WorkingStep.WorkingStepCode;
            txtDes.Text = WorkingStep.Description;
        }

        private void frmWorkingDetail_Load(object sender, EventArgs e)
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
        void saveData()
        {
            if (!ValidateForm())
                return;

            WorkingStep.WorkingStepCode = txtCode.Text.Trim();
            WorkingStep.SortOrder = (int)txtSortOrder.Value;
            WorkingStep.Description = txtDes.Text.Trim();
            WorkingStep.UpdatedBy = Global.AppUserName;
            WorkingStep.UpdatedDate = DateTime.Now;

            if (WorkingStep.ID <= 0)
            {
                WorkingStep.CreatedBy = Global.AppUserName;
                WorkingStep.CreatedDate = DateTime.Now;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveData();
            this.DialogResult = DialogResult.OK;
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            saveData();
            WorkingStep = new WorkingStepModel();
            loadData();
        }

        private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSave_Click(null, null);
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
