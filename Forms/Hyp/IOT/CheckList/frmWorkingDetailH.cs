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
    public partial class frmWorkingDetailH : _Forms
    {
        public WorkingModel WorkingModel = new WorkingModel();
        public int GroupID;
        public int StepID;
        public frmWorkingDetailH()
        {
            InitializeComponent();
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
                dt = LibIE.Select(string.Format("SELECT * FROM dbo.WorkingStep WHERE ProductGroupID = {0} ORDER BY SortOrder", groupID));
            }
            cboStep.Properties.DataSource = dt;
            cboStep.Properties.DisplayMember = "WorkingStepCode";
            cboStep.Properties.ValueMember = "ID";
        }

        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            if (WorkingModel.ID == 0)
            {
                txtSortOrder.Value = 1;
            }
            
            txtMax.Value = WorkingModel.MaxValue;
            txtMin.Value = WorkingModel.MinValue;
            txtName.Text = WorkingModel.WorkingName;
            txtPeriodValue.Text = WorkingModel.PeriodValue;
            txtSortOrder.Value = WorkingModel.SortOrder;
            txtUnit.Text = WorkingModel.Unit;

            cboValueType.SelectedIndex = WorkingModel.ValueType;
            cboComPort.SelectedIndex = WorkingModel.Comport - 1;
            
            cboStep.EditValue = WorkingModel.WorkingStepID;
            cboGroup.EditValue = WorkingModel.ProductGroupID;
            //chkAutoValue.Checked = WorkingModel.IsGetAutoValue;

            cboComPort.Enabled = chkAutoValue.Checked = WorkingModel.IsGetAutoValueComport;
            chkIP.Checked = txtIP.Enabled = txtPort.Enabled = WorkingModel.IsGetAutoValueIP;
            txtPort.Text = WorkingModel.Port;
            txtIP.Text = WorkingModel.IpAddress;
            cboCheckValueType.SelectedIndex = WorkingModel.CheckValueType;

        }

        private void frmWorkingDetail_Load(object sender, EventArgs e)
        {
           
            loadGroup();            
            loadData();
            if (WorkingModel.ID == 0)
            {
                cboGroup.EditValue = GroupID;
                cboStep.EditValue = StepID;
            }
            txtName.Focus();
        }

        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (cboGroup.EditValue == null)
            {
                MessageBox.Show("Xin hãy chọn nhóm sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            if (cboStep.EditValue == null)
            {
                MessageBox.Show("Xin hãy chọn công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (cboValueType.SelectedIndex < 0)
            {
                MessageBox.Show("Xin hãy chọn kiểu kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            //if (cboCheckValueType.SelectedIndex < 1)
            //{
            //    MessageBox.Show("Xin hãy chọn loại giá trị kiểm tra.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return false;
            //}

            if (txtName.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền Mô tả công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            string sql1 = "";
            if (WorkingModel.ID > 0)
            {
                 sql1 = $"SELECT ID from Working WHERE SortOrder = '{txtSortOrder.Value}' and ID <>'{WorkingModel.ID}' and WorkingStepID={WorkingModel.WorkingStepID} ";
              
            }
            else
            { 
                 sql1 = $"SELECT ID from Working WHERE SortOrder = '{txtSortOrder.Value}' and WorkingStepID={StepID}";
            }
            int id = TextUtils.ToInt(LibIE.ExcuteScalar(sql1));
            if (id > 0)
            {
                MessageBox.Show("Trùng STT", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
               // return false;
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

            WorkingModel.ProductGroupID = TextUtils.ToInt(cboGroup.EditValue); 
            WorkingModel.WorkingStepID = TextUtils.ToInt(cboStep.EditValue);
            WorkingModel.WorkingName = txtName.Text.Trim();
            WorkingModel.SortOrder = (int)txtSortOrder.Value;

            WorkingModel.Unit = txtUnit.Text.Trim();
            WorkingModel.ValueType = cboValueType.SelectedIndex;
            WorkingModel.ValueTypeName = cboValueType.SelectedIndex == 0 ? "Check mark" : "Giá trị\n数値";
            WorkingModel.MaxValue = TextUtils.ToDecimal(txtMax.Value);
            WorkingModel.MinValue = TextUtils.ToDecimal(txtMin.Value);
           // WorkingModel.PeriodValue = WorkingModel.MinValue + "~" + WorkingModel.MaxValue;

            WorkingModel.IsGetAutoValueComport = chkAutoValue.Checked;
            WorkingModel.Comport = cboComPort.SelectedIndex + 1;

            WorkingModel.CheckValueType = cboCheckValueType.SelectedIndex;
            WorkingModel.IsGetAutoValueIP = chkIP.Checked;
            WorkingModel.IpAddress = txtIP.Text.Trim();
            WorkingModel.Port = txtPort.Text.Trim();

            if (WorkingModel.ValueType == 0)//checkmark
            {
                WorkingModel.PeriodValue = txtPeriodValue.Text.Trim();
            }
            else
            {
                if (WorkingModel.CheckValueType == 1)//giá trị dạng số
                {
                    if (WorkingModel.MinValue == WorkingModel.MaxValue)
                    {
                        WorkingModel.PeriodValue = WorkingModel.MaxValue.ToString("n3");
                    }
                    else
                    {
                        WorkingModel.PeriodValue = WorkingModel.MinValue.ToString("n3") + "~" + WorkingModel.MaxValue.ToString("n3");
                    }
                }
                else
                {
                    WorkingModel.PeriodValue = txtPeriodValue.Text.Trim();
                }
            }

            if (WorkingModel.ID > 0)
            {
                WorkingBO.Instance.Update(WorkingModel);
                LibIE.ExcuteSQL(string.Format("EXEC spUpdateMultiProductWorking {0}, 1", WorkingModel.ID));
            }
            else
            {
                WorkingModel.ID = (int)WorkingBO.Instance.Insert(WorkingModel);
                LibIE.ExcuteSQL(string.Format("EXEC CreateProductWorking_ByWorkingID {0}", WorkingModel.ID));
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
                WorkingModel = new WorkingModel();
                loadData();
                txtName.Focus();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkAutoValue_CheckedChanged(object sender, EventArgs e)
        {
            cboComPort.Enabled = chkAutoValue.Checked;
            if (!chkAutoValue.Checked)
            {
                cboComPort.SelectedIndex = -1;
            }
        }

        private void cboGroup_EditValueChanged(object sender, EventArgs e)
        {
            loadStep();
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

        private void chkIP_CheckedChanged(object sender, EventArgs e)
        {
            txtIP.Enabled = txtPort.Enabled = chkIP.Checked;
            if (!chkIP.Checked)
            {
                txtIP.Text = txtPort.Text = "";
            }
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
