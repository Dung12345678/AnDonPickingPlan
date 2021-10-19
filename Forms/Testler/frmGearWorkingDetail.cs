using System;
using BMS;
using BMS.Business;
using BMS.Model;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmGearWorkingDetail : _Forms
    {
        public int GearID = 0;
        public GearWorkingModel GearWorking = new GearWorkingModel();
        public frmGearWorkingDetail()
        {
            InitializeComponent();
        }
        private bool ValidateForm()
        {
            //if (cboStep.EditValue == null)
            //{
            //    MessageBox.Show("Xin hãy chọn một công đoạn.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return false;
            //}

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

            return true;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (SaveData())
                this.DialogResult = DialogResult.OK;
        }
        bool SaveData()
        {
            if (!ValidateForm())
                return false;

            GearWorking.GearID = GearID;
            GearWorking.WorkingName = txtName.Text.Trim();
            GearWorking.SortOrder = (int)txtSortOrder.Value;

            GearWorking.Unit = txtUnit.Text;
            GearWorking.MaxValue = TextUtils.ToDecimal(txtMax.Value);
            GearWorking.MinValue = TextUtils.ToDecimal(txtMin.Value);
            GearWorking.DefaultValue = TextUtils.ToDecimal(numDefautValue.Value);
            GearWorking.TanSuat = txtTanSuat.Text;
            GearWorking.ValueType = TextUtils.ToInt(cboValueType.SelectedIndex);
            GearWorking.CheckValueType = TextUtils.ToInt(cboCheckValueType.SelectedIndex);

            if (GearWorking.ValueType == 0) //check mark (OK vs NG)
            {
                GearWorking.PeriodValue = txtPeriodValue.Text.Trim();             
            }
            else
            {
                if (GearWorking.CheckValueType == 1)
                {
                    if (GearWorking.MinValue == GearWorking.MaxValue)
                    {
                        GearWorking.PeriodValue = GearWorking.MaxValue.ToString("n3");
                    }
                    else
                    {
                        GearWorking.PeriodValue = GearWorking.MinValue.ToString("n3") + "~" + GearWorking.MaxValue.ToString("n3");
                    }
                    
                }
                else
                {
                    GearWorking.PeriodValue = txtPeriodValue.Text.Trim();
                }
            }
            GearWorking.ValueTypeName = cboCheckValueType.SelectedItem.ToString();
            if (GearWorking.ID > 0)
            {
                GearWorkingBO.Instance.Update(GearWorking);
            }
            else
            {
                GearWorking.ID = (int)GearWorkingBO.Instance.Insert(GearWorking);
            }
            return true;
        }
        void LoadData()
        {
            if (GearWorking.ID == 0)
            {
                txtSortOrder.Value = 0;
                cboValueType.SelectedIndex = 1;
                cboCheckValueType.SelectedIndex = 1;
            }
            else
            {
                txtMax.Value = GearWorking.MaxValue;
                txtMin.Value = GearWorking.MinValue;
                numDefautValue.Value = GearWorking.DefaultValue;
                txtTanSuat.Text = GearWorking.TanSuat;
                txtName.Text = GearWorking.WorkingName;
                txtPeriodValue.Text = GearWorking.PeriodValue;
                txtSortOrder.Value = GearWorking.SortOrder;
                txtUnit.Text = GearWorking.Unit;
                cboValueType.SelectedIndex = GearWorking.ValueType;
                cboCheckValueType.SelectedIndex = GearWorking.CheckValueType;
            }
        }

        private void frmGearWorkingDetail_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            if (SaveData())
            {
                GearWorking = new GearWorkingModel();
                LoadData();
            }
        }

        private void txtMin_ValueChanged(object sender, EventArgs e)
        {
            txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
        }

        private void txtMax_ValueChanged(object sender, EventArgs e)
        {
            txtPeriodValue.Text = txtMin.Value + "~" + txtMax.Value;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmGearWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
