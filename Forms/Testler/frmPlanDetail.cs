using BMS.Business;
using BMS.Model;
using System;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmPlanDetail : _Forms
    {
        public PlanOrderModel PlanOrderModel = new PlanOrderModel();
        public frmPlanDetail()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {            
            txtOrderCode.Text = PlanOrderModel.OrderCode;
            txtGearCode.Text = PlanOrderModel.GearCode;
            txtQty.Value = PlanOrderModel.Qty;
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
            if (txtOrderCode.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền HYP.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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

            PlanOrderModel.OrderCode = txtOrderCode.Text;
            PlanOrderModel.GearCode = txtGearCode.Text;
            PlanOrderModel.Qty = TextUtils.ToInt(txtQty.Value);

            if (PlanOrderModel.ID <= 0)
            {
                GearBO.Instance.Insert(PlanOrderModel);
            }
            else
            {
                GearBO.Instance.Update(PlanOrderModel);
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
                PlanOrderModel = new PlanOrderModel();
                loadData();
            }
        }

        private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
        }
    }
}
