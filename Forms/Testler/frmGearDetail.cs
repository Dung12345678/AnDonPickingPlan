using BMS.Business;
using BMS.Model;
using System;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmGearDetail : _Forms
    {
        public GearModel GearModel = new GearModel();

        public int ProductID = 0;

        public frmGearDetail()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {            
            txtHYP.Text = GearModel.HYP;
            txtGear1.Text = GearModel.Gear1;
            txtGear2.Text = GearModel.Gear2;
            txtGear3.Text = GearModel.Gear3;
            txtLap.Text = GearModel.Lap;

            txtLotSize.Value = GearModel.LotSize;
            txtGThieu.Value = GearModel.GThieu;
            txtGGiaCongThem.Value = GearModel.GGiaCongThem;
            txtWipG.Value = GearModel.WipG;
            txtOnhandG.Value = GearModel.OnhandG;
            txtStockGCQT.Value = GearModel.StockGCQT;
            txtWipGCQT.Value = GearModel.WipGCQT;
            txtGCQTThieu.Value = GearModel.GCQTThieu;
            txtWipTruc.Value = GearModel.WipTruc;
            txtOrderTrucMoi.Value = GearModel.OrderTrucMoi;
            txtHypNeed.Value = GearModel.HypNeed;
            nmrQtyImage.Value = GearModel.QtyImage;
            nmrQtyAudio.Value = GearModel.QtyAudio;
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
            if (txtHYP.Text.Trim() == "")
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

            GearModel.HYP = txtHYP.Text;
            GearModel.Gear1 = txtGear1.Text;
            GearModel.Gear2 = txtGear2.Text;
            GearModel.Gear3 = txtGear3.Text;
            GearModel.Lap = txtLap.Text;
            GearModel.QtyImage = TextUtils.ToInt(nmrQtyImage.Value);
            GearModel.QtyAudio = TextUtils.ToInt(nmrQtyAudio.Value);
            

            GearModel.LotSize = TextUtils.ToInt(txtLotSize.Value);
            GearModel.GThieu = TextUtils.ToInt(txtGThieu.Value);
            GearModel.GGiaCongThem = TextUtils.ToInt(txtGGiaCongThem.Value);
            GearModel.WipG = TextUtils.ToInt(txtWipG.Value);
            GearModel.OnhandG = TextUtils.ToInt(txtOnhandG.Value);
            GearModel.StockGCQT = TextUtils.ToInt(txtStockGCQT.Value);
            GearModel.WipGCQT = TextUtils.ToInt(txtWipGCQT.Value);
            GearModel.GCQTThieu = TextUtils.ToInt(txtGCQTThieu.Value);
            GearModel.WipTruc = TextUtils.ToInt(txtWipTruc.Value);
            GearModel.OrderTrucMoi = TextUtils.ToInt(txtOrderTrucMoi.Value);
            GearModel.HypNeed = TextUtils.ToInt(txtHypNeed.Value);

            if (GearModel.ID <= 0)
            {
                GearBO.Instance.Insert(GearModel);
            }
            else
            {
                GearBO.Instance.Update(GearModel);
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
                GearModel = new GearModel();
                loadData();
            }
        }

        private void frmWorkingDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.DialogResult = DialogResult.OK;
        }
    }
}
