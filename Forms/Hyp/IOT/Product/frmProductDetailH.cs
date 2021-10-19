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
	public partial class frmProductDetailH : _Forms
	{
		public ProductModel productModel = new ProductModel();
		public int ProductGroupID { get; set; }

		public frmProductDetailH()
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
			txtCoilCode.Text = productModel.CoilCode;
			txtMotorCode.Text = productModel.MotorCode;
			dtpDate.EditValue = productModel.ProductDate;
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
			txtJobNumber.Value = productModel.JobNumber;
			txtShootNumber.Value = productModel.ShootNumber;
			txtBuLongNumber.Value = productModel.BuLongNumber;
			txtSetForce.Value = productModel.SetForce;
			txtXuatLucSealCap.Text = productModel.XuatLucSealCap;
			txtLoaiMoNapDau.Text = productModel.LoaiMoNapDau;
			txtLoaiMoBomHopSo.Text = productModel.LoaiMoBomHopSo;
			txtLucCheckMin.Value = productModel.LucCheckGearMin;
			txtLucCheckMax.Value = productModel.LucCheckGearMax;

			txtEngineWeightMax.Value = TextUtils.ToDecimal(productModel.EngineWeightMax);
			txtEngineWeightMin.Value= TextUtils.ToDecimal(productModel.EngineWeightMin);
			txtEngineWeight.Text = TextUtils.ToString(productModel.EngineWeight);
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
					dt = LibIE.Select("select top 1 ProductCode from Product where ProductCode = '" + txtCode.Text.Trim() + "' and ID <> " + productModel.ID);
				}
				else
				{
					dt = LibIE.Select("select top 1 ProductCode from Product where ProductCode = '" + txtCode.Text.Trim() + "'");
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

			productModel.ProductGroupID = TextUtils.ToInt(cboGroup.EditValue);//"NHÓM SẢN PHẨM"
			productModel.ProductTypeCode = txtProductTypeCode.Text.Trim();//"LOẠI SẢN PHẨM#"
			productModel.ProductCode = txtCode.Text.Trim(); //Product code
			productModel.ProductName = txtDes.Text.Trim();//MÔ TẢ SẢN PHẨM
			productModel.RatioCode = txtRatio.Text.Trim();     //Giảm tốc   
			productModel.ProductDate = TextUtils.ToDate2(dtpDate.EditValue);//NGÀY LẬP
			productModel.XuatLucSealCap = txtXuatLucSealCap.Text.Trim();//"HƯỚNG TRỤC XUẤT LỰC SEAL CAP"
			productModel.LoaiMoNapDau = txtLoaiMoNapDau.Text.Trim();//LOẠI MỠ NẮP DẦU
			productModel.LoaiMoBomHopSo = txtLoaiMoBomHopSo.Text.Trim();//LOẠI MỠ BƠM HỘP SỐ
																   //"LỰC CHECK 1  NẮP - HỘP (N.m)"
			productModel.LucCheckGearMin = txtLucCheckMin.Value;
			productModel.LucCheckGearMax = txtLucCheckMax.Value;
			productModel.GunNumber = TextUtils.ToInt(txtGunNumber.Value); //SÚNG
			productModel.JobNumber = TextUtils.ToInt(txtJobNumber.Value);//JOB
			productModel.QtyOcBanGa = TextUtils.ToInt(txtQtyOcBanGa.Value);   //BẮN GÁ
			productModel.BuLongNumber = TextUtils.ToInt(txtBuLongNumber.Value);// SỐ LUỢNG BULONG
			productModel.ShootNumber = TextUtils.ToInt(txtShootNumber.Value);//SỐ LẦN BẮN
			productModel.SetForce = TextUtils.ToInt(txtSetForce.Value);// đặt lực
			productModel.MotorCode = txtMotorCode.Text.Trim();    //MÃ MOTOR
			productModel.CoilCode = txtCoilCode.Text.Trim();// Mã Coil
			productModel.DienAp = txtDienAp.Value;//Điện áp (V)
			productModel.TanSo = txtTanSo.Value;//Tần số (Hz)
										   //Dòng điện Io
			productModel.DongDienMin = txtDongDienMin.Value;
			productModel.DongDienMax = txtDongDienMax.Value;
			//Nhập lực Wo
			productModel.NhapLucMax = txtNhapLucMax.Value;
			productModel.NhapLucMin = txtNhapLucMin.Value;
			//Vòng quay
			productModel.VongQuayMax = txtVongQuayMax.Value;
			productModel.VongQuayMin = txtVongQuayMin.Value;
			// Tiếng ồn dB
			productModel.TiengOnMax = txtTiengOnMax.Value;
			productModel.TiengOnMin = txtTiengOnMin.Value;
			//Độ rung FFT
			productModel.DoRungFFTMax = txtDoRungMax.Value;
			productModel.DoRungFFTMin = txtDoRungMin.Value;

			// Trọng lượng sản phẩm khi có mỡ và đã lắp hoàn chỉnh
			productModel.EngineWeightMax = TextUtils.ToDouble(txtEngineWeightMax.Value);
			productModel.EngineWeightMin = TextUtils.ToDouble(txtEngineWeightMin.Value);
			productModel.EngineWeight = TextUtils.ToDouble(txtEngineWeight.Text.Trim());

			productModel.Goal = txtGoal.Text.Trim();// đích  
			if (productModel.ID > 0)
			{
				WorkingBO.Instance.Update(productModel);
			}
			else
			{
				productModel.ID = (int)WorkingBO.Instance.Insert(productModel);
				//Sinh các công đoạn và mục kiểm tra của sản phẩm theo định dạng của nhóm sản phẩm
				LibIE.ExcuteSQL(string.Format("EXEC dbo.CreateProductWorking_ByGroupID {0},{1}", productModel.ProductGroupID, productModel.ID));
			}
			/*
             * Cập nhật lại các mục check sản phẩm theo các rule quy định của các công đoạn
             */
			LibIE.ExcuteSQL(string.Format("EXEC spUpdateProductWorking_ByProductID_New {0},{1}", productModel.ID, productModel.MURATA == "3" ? 0 : 1));
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
				productModel = new ProductModel();
				loadData();
			}
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
