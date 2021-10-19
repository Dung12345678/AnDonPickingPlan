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
    public partial class frmProductDetail : _Forms
    {
        public ProductModel productModel = new ProductModel();
        public int ProductGroupID { get; set; }

        public frmProductDetail()
        {
            InitializeComponent();
        }
        void loadGroup()
        {
            DataTable dt = TextUtils.Select("SELECT * FROM dbo.ProductGroup WITH(NOLOCK)");
            cboGroup.Properties.DataSource = dt;
            cboGroup.Properties.DisplayMember = "ProductGroupName";
            cboGroup.Properties.ValueMember = "ID";
        }


        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            cboGroup.EditValue = productModel.ProductGroupID;
            txtProductTypeCode.Text = productModel.ProductTypeCode;
            txtCode.Text = productModel.ProductCode;
            txtDes.Text = productModel.ProductName;
            txtRatio.Text = productModel.RatioCode;
            txtPinNumber.Text = productModel.PinNumber;
            txtPinNumberIN.Text = productModel.PinNumberIN;
            txtCoilCode.Text = productModel.CoilCode;
            txtMotorCode.Text = productModel.MotorCode;
            txtLoaiMo.Text = productModel.LoaiMo;
            txtLuongMo.Value = productModel.LuongMo;
            txtMurata.Text = productModel.MURATA;
            txtUnitMotor.Text = productModel.UnitMotor;
            txtMauSon.Text = productModel.MauSon;
            dtpDate.EditValue = productModel.ProductDate;
            dtpMEDEDAY.EditValue = productModel.MEDEDAY;
            txtHuongHopCau.Text = productModel.HuongHopCau;

            txtDienAp.Value = productModel.DienAp;
            txtTanSo.Value = productModel.TanSo;
            txtDongDienMax.Value = productModel.DongDienMax;
            txtDongDienMin.Value = productModel.DongDienMin;
            txtNhapLucMax.Value = productModel.NhapLucMax;
            txtNhapLucMin.Value = productModel.NhapLucMin;
            txtVongQuayMax.Value = productModel.VongQuayMax;
            txtVongQuayMin.Value = productModel.VongQuayMin;
            txtTiengOnMax.Value = productModel.TiengOnMax;
            txtTiengOnMin.Value = productModel.TiengOnMin;
            txtDoRungMax.Value = productModel.DoRungFFTMax;
            txtDoRungMin.Value = productModel.DoRungFFTMin;

            txtGunNumber.Value = productModel.GunNumber;
            txtQtyOcBanGa.Value = productModel.QtyOcBanGa;
            txtQtyOcBanThat.Value = productModel.QtyOcBanThat;
            txtJobNumber.Value = productModel.JobNumber;
            txtShootNumber.Value = productModel.ShootNumber;
            txtBuLongNumber.Value = productModel.BuLongNumber;
            txtSetForce.Value = productModel.SetForce;

            txtDifferenceMax.Value = productModel.DifferenceMax;
            txtDifferenceMin.Value = productModel.DifferenceMin;
            txtSizeSSSMax.Value = productModel.SizeSSSMax;
            txtSizeSSSMin.Value = productModel.SizeSSSMin;
            txtSpaceCenterMax.Value = productModel.SpaceCenterMax;
            txtSpaceCenterMin.Value = productModel.SpaceCenterMin;
            txtSpaceCycloidDiskMin.Value = productModel.SpaceCycloidDiskMin;
            txtSpaceCycloidDiskMax.Value = productModel.SpaceCycloidDiskMax;

            txtLucCheckGearMax.Value = productModel.LucCheckGearMax;
            txtLucCheckGearMin.Value = productModel.LucCheckGearMin;
            txtLucCheckGearMotorMax.Value = productModel.LucCheckGearMotorMax;
            txtLucCheckGearMotorMin.Value = productModel.LucCheckGearMotorMin;
            txtLucCheck3Max.Value = productModel.LucCheck3Max;//Lực check cụm trục tốc độ thấp 511-512
            txtLucCheck3Min.Value = productModel.LucCheck3Min;
            txtLucCheck4Max.Value = productModel.LucCheck4Max;//Lực check phần xuất lực Naknishi
            txtLucCheck4Min.Value = productModel.LucCheck4Min;
            txtLucCheck5Max.Value = productModel.LucCheck5Max;//Lực check Ốc chặn mỡ Nakanishi, 6000
            txtLucCheck5Min.Value = productModel.LucCheck5Min;

            txtDoDao1Max.Value = productModel.DoDao1Max;//ĐẢO TRƯỚC ÉP VÒNG BI
            txtDoDao1Min.Value = productModel.DoDao1Min;
            txtDoDao2Max.Value = productModel.DoDao2Max;//ĐẢO SAU ÉP VÒNG BI
            txtDoDao2Min.Value = productModel.DoDao2Min;
            txtDoDao3Max.Value = productModel.DoDao3Max;
            txtDoDao3Min.Value = productModel.DoDao3Min;
            txtDoDao4Max.Value = productModel.DoDao4Max;
            txtDoDao4Min.Value = productModel.DoDao4Min;

            txtLuongMo1Max.Value = productModel.LuongMo1Max;//Mỡ cho trục tốc độ thấp
            txtLuongMo1Min.Value = productModel.LuongMo1Min;//
            txtLuongMo2Max.Value = productModel.LuongMo2Max;//Lượng mỡ cho hộp
            txtLuongMo2Min.Value = productModel.LuongMo2Min;//
            txtLuongMo3Max.Value = productModel.LuongMo3Max;//Mỡ cho vòng bi chịu tải & Carrier SSS
            txtLuongMo3Min.Value = productModel.LuongMo3Min;//
            txtLuongMo4Max.Value = productModel.LuongMo4Max;//Lượng mỡ cho vòng bi không tải
            txtLuongMo4Min.Value = productModel.LuongMo4Min;//
            txtLuongMo5Max.Value = productModel.LuongMo5Max;//Mỡ cho Carrier
            txtLuongMo5Min.Value = productModel.LuongMo5Min;//
            txtLuongMo6Max.Value = productModel.LuongMo6Max;//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512
            txtLuongMo6Min.Value = productModel.LuongMo6Min;//
            txtLuongMo7Max.Value = productModel.LuongMo7Max;//
            txtLuongMo7Min.Value = productModel.LuongMo7Min;//
        
            //nvthao thêm ngày 03.01.2020
            //Thêm trường Đích lưu cho sản phẩm
            txtGoal.Text = productModel.Goal;
        }

        private void frmWorkingDetail_Load(object sender, EventArgs e)
        {
            loadGroup();
            loadData();

            if (ProductGroupID > 0)
            {
                cboGroup.EditValue = ProductGroupID;
            }
        }

        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (cboGroup.EditValue == null)
            {
                MessageBox.Show("Xin hãy chọn một nhóm sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (string.IsNullOrEmpty(txtProductTypeCode.Text.Trim()))
            {
                MessageBox.Show("Xin hãy nhập loại sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (txtCode.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền mã sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (productModel.ID > 0)
                {                    
                    dt = TextUtils.Select("select top 1 ProductCode from Product where ProductCode = '" + txtCode.Text.Trim() + "' and ID <> " + productModel.ID);
                }
                else
                {
                    dt = TextUtils.Select("select top 1 ProductCode from Product where ProductCode = '" + txtCode.Text.Trim() + "'");
                }
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Mã này đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                }
            }
            if (txtDes.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền tên sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
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

            productModel.ProductGroupID = TextUtils.ToInt(cboGroup.EditValue);
            productModel.ProductTypeCode = txtProductTypeCode.Text.Trim();
            productModel.ProductCode = txtCode.Text.Trim();
            productModel.ProductName = txtDes.Text.Trim();
            productModel.RatioCode = txtRatio.Text.Trim();
            productModel.PinNumber = txtPinNumber.Text.Trim();
            productModel.PinNumberIN = txtPinNumberIN.Text.Trim();
            productModel.CoilCode = txtCoilCode.Text.Trim();
            productModel.MotorCode = txtMotorCode.Text.Trim();
            productModel.LoaiMo = txtLoaiMo.Text.Trim();
            productModel.LuongMo = txtLuongMo.Value;
            productModel.MURATA = txtMurata.Text.Trim();
            productModel.UnitMotor = txtUnitMotor.Text.Trim();
            productModel.MauSon = txtMauSon.Text.Trim();
            productModel.MEDEDAY = TextUtils.ToDate2(dtpMEDEDAY.EditValue);
            productModel.ProductDate = TextUtils.ToDate2(dtpDate.EditValue);
            productModel.HuongHopCau = txtHuongHopCau.Text.Trim();

            productModel.DienAp = txtDienAp.Value;
            productModel.TanSo = txtTanSo.Value;
            productModel.DongDienMin = txtDongDienMin.Value;
            productModel.DongDienMax = txtDongDienMax.Value;
            productModel.NhapLucMax = txtNhapLucMax.Value;
            productModel.NhapLucMin = txtNhapLucMin.Value;
            productModel.VongQuayMax = txtVongQuayMax.Value;
            productModel.VongQuayMin = txtVongQuayMin.Value;
            productModel.TiengOnMax = txtTiengOnMax.Value;
            productModel.TiengOnMin = txtTiengOnMin.Value;
            productModel.DoRungFFTMax = txtDoRungMax.Value;
            productModel.DoRungFFTMin = txtDoRungMin.Value;

            productModel.GunNumber = TextUtils.ToInt(txtGunNumber.Value);
            productModel.JobNumber = TextUtils.ToInt(txtJobNumber.Value);
            productModel.QtyOcBanGa = TextUtils.ToInt(txtQtyOcBanGa.Value);
            productModel.QtyOcBanThat = TextUtils.ToInt(txtQtyOcBanThat.Value);
            productModel.BuLongNumber = TextUtils.ToInt(txtBuLongNumber.Value);
            productModel.ShootNumber = TextUtils.ToInt(txtShootNumber.Value);
            productModel.SetForce = TextUtils.ToInt(txtSetForce.Value);

            productModel.DifferenceMax = txtDifferenceMax.Value;
            productModel.DifferenceMin = txtDifferenceMin.Value;
            productModel.SizeSSSMax = txtSizeSSSMax.Value;
            productModel.SizeSSSMin = txtSizeSSSMin.Value;
            productModel.SpaceCenterMax = txtSpaceCenterMax.Value;
            productModel.SpaceCenterMin = txtSpaceCenterMin.Value;
            productModel.SpaceCycloidDiskMin = txtSpaceCycloidDiskMin.Value;
            productModel.SpaceCycloidDiskMax = txtSpaceCycloidDiskMax.Value;

            productModel.LucCheckGearMax = txtLucCheckGearMax.Value;
            productModel.LucCheckGearMin = txtLucCheckGearMin.Value;
            productModel.LucCheckGearMotorMax = txtLucCheckGearMotorMax.Value;
            productModel.LucCheckGearMotorMin = txtLucCheckGearMotorMin.Value;
            productModel.LucCheck3Max = txtLucCheck3Max.Value;//Lực check cụm trục tốc độ thấp 511-512
            productModel.LucCheck3Min = txtLucCheck3Min.Value;
            productModel.LucCheck4Max = txtLucCheck4Max.Value;//Lực check phần xuất lực Naknishi
            productModel.LucCheck4Min = txtLucCheck4Min.Value;
            productModel.LucCheck5Max = txtLucCheck5Max.Value;//Lực check Ốc chặn mỡ Nakanishi, 6000
            productModel.LucCheck5Min = txtLucCheck5Min.Value;

            productModel.DoDao1Max = txtDoDao1Max.Value;//ĐẢO TRƯỚC ÉP VÒNG BI
            productModel.DoDao1Min = txtDoDao1Min.Value;
            productModel.DoDao2Max = txtDoDao2Max.Value;//ĐẢO SAU ÉP VÒNG BI
            productModel.DoDao2Min = txtDoDao2Min.Value;
            productModel.DoDao3Max = txtDoDao3Max.Value;
            productModel.DoDao3Min = txtDoDao3Min.Value;
            productModel.DoDao4Max = txtDoDao4Max.Value;
            productModel.DoDao4Min = txtDoDao4Min.Value;

            productModel.LuongMo1Max = txtLuongMo1Max.Value;//Mỡ cho trục tốc độ thấp
            productModel.LuongMo1Min = txtLuongMo1Min.Value;//
            productModel.LuongMo2Max = txtLuongMo2Max.Value;//Lượng mỡ cho hộp
            productModel.LuongMo2Min = txtLuongMo2Min.Value;//
            productModel.LuongMo3Max = txtLuongMo3Max.Value;//Mỡ cho vòng bi chịu tải & Carrier SSS
            productModel.LuongMo3Min = txtLuongMo3Min.Value;//
            productModel.LuongMo4Max = txtLuongMo4Max.Value;//Lượng mỡ cho vòng bi không tải
            productModel.LuongMo4Min = txtLuongMo4Min.Value;//
            productModel.LuongMo5Max = txtLuongMo5Max.Value;//Mỡ cho Carrier
            productModel.LuongMo5Min = txtLuongMo5Min.Value;//
            productModel.LuongMo6Max = txtLuongMo6Max.Value;//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512
            productModel.LuongMo6Min = txtLuongMo6Min.Value;//
            productModel.LuongMo7Max = txtLuongMo7Max.Value;//
            productModel.LuongMo7Min = txtLuongMo7Min.Value;//

            productModel.Goal = txtGoal.Text.Trim();

            if (productModel.ID > 0)
            {
                WorkingBO.Instance.Update(productModel);
            }
            else
            {
                productModel.ID = (int)WorkingBO.Instance.Insert(productModel);
                //Sinh các công đoạn và mục kiểm tra của sản phẩm theo định dạng của nhóm sản phẩm
                TextUtils.ExcuteSQL(string.Format("EXEC dbo.CreateProductWorking_ByGroupID {0},{1}", productModel.ProductGroupID, productModel.ID));
            }

            /*
             * Cập nhật lại các mục check sản phẩm theo các rule quy định của các công đoạn
             */
            TextUtils.ExcuteSQL(string.Format("EXEC spUpdateProductWorking_ByProductID_New {0},{1}", productModel.ID, productModel.MURATA == "3" ? 0 : 1));
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
                productModel = new ProductModel();
                loadData();
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

        private void control_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
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
