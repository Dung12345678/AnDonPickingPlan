using IE.Business;
using IE.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmWorkingStepDetailH : _Forms
    {
        public WorkingStepModel WorkingStep = new WorkingStepModel();

        public ArrayList ArrStep { get; set; }

        public frmWorkingStepDetailH()
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
            if (txtSortOrder.Value <= 0)
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
            var lst = ArrStep.ToArray().ToList();
            if (WorkingStep.ID > 0)
            {
                lst = lst.Where(o => ((WorkingStepModel)o).ID != WorkingStep.ID && ((WorkingStepModel)o).WorkingStepCode == txtCode.Text.Trim()).ToList();
            }
            else
            {
                lst = lst.Where(o => ((WorkingStepModel)o).WorkingStepCode == txtCode.Text.Trim()).ToList();
            }
            if (lst != null && lst.Count > 0)
            {
                MessageBox.Show("Mã công đoạn này đã tồn tại.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
            //bool isError = false;
            if (!ValidateForm())
            {
                return false;
            }

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
            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (saveData())
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
