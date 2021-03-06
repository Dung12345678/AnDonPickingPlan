
using ST.Business;
using ST.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmAddEditPlan : _Forms
	{
		public SonPlanModel sonPlanModel = new SonPlanModel();
		private int type;

		public int Type
		{
			get
			{
				return type;
			}

			set
			{
				type = value;
			}
		}

		public frmAddEditPlan(int typeEvent)
		{
			InitializeComponent();
			Type = typeEvent;
			switch (typeEvent)
			{
				case 1:
					this.Text = "Thêm kế hoạch";
					break;
				case 2:
					this.Text = "Sửa kế hoạch";
					break;
			}
		}

		void LoadDataToForm()
		{
			txbPartCode.Text = sonPlanModel.PartCode;
			txbOrderCode.Text = sonPlanModel.OrderCode;
			txbNote.Text = sonPlanModel.Note;
			txbConfirmCode.Text = sonPlanModel.ConfirmCode;
			txbLotSize.Value = sonPlanModel.LotSize;
			txbNG.Value = sonPlanModel.NG;
			txbOP.Value = sonPlanModel.OP;
			txtCNT.Value = sonPlanModel.Cnt;
			txbQtyPlan.Value = sonPlanModel.QtyPlan;
			txbRealProdQty.Value = sonPlanModel.RealProdQty;
			txbSaleCode.Text = sonPlanModel.SaleCode;
			txbShipTo.Text = sonPlanModel.ShipTo;
			txbShipVia.Text = sonPlanModel.ShipVia;
			txbWorkerCode.Text = sonPlanModel.WorkerCode;
			dtpDateExported.Value = TextUtilsStock.ToDate2(sonPlanModel.DateExported) == null ? DateTime.Now : TextUtilsStock.ToDate3(sonPlanModel.DateExported);
			dtpPrintedDate.Value = TextUtilsStock.ToDate2(sonPlanModel.PrintedDate) == null ? DateTime.Now : TextUtilsStock.ToDate3(sonPlanModel.PrintedDate);
			dtpProdDate.Value = TextUtilsStock.ToDate2(sonPlanModel.ProdDate) == null ? DateTime.Now : TextUtilsStock.ToDate3(sonPlanModel.ProdDate);
		}

		private bool ValidateForm()
		{
			if (txbPartCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập mã linh kiện.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txbOrderCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập mã order.", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;

		}

		bool SaveData()
		{
			if (!ValidateForm())
			{
				return false;
			}

			sonPlanModel.PartCode = txbPartCode.Text;
			sonPlanModel.OrderCode = txbOrderCode.Text;
			sonPlanModel.Note = txbNote.Text;
			sonPlanModel.ConfirmCode = txbConfirmCode.Text;
			sonPlanModel.LotSize = TextUtilsStock.ToInt(txbLotSize.Value);
			sonPlanModel.NG = TextUtilsStock.ToInt(txbNG.Value);
			sonPlanModel.OP = TextUtilsStock.ToInt(txbOP.Value);
			sonPlanModel.Cnt = TextUtilsStock.ToInt(txtCNT.Value);
			sonPlanModel.QtyPlan = TextUtilsStock.ToInt(txbQtyPlan.Value);
			sonPlanModel.RealProdQty = TextUtilsStock.ToInt(txbRealProdQty.Value);
			sonPlanModel.SaleCode = txbSaleCode.Text;
			sonPlanModel.ShipTo = txbShipTo.Text;
			sonPlanModel.ShipVia = txbShipVia.Text;
			sonPlanModel.WorkerCode = txbWorkerCode.Text;
			sonPlanModel.DateExported = TextUtilsStock.ToDate2(dtpDateExported.Value);
			sonPlanModel.PrintedDate = TextUtilsStock.ToDate2(dtpPrintedDate.Value);
			sonPlanModel.ProdDate = TextUtilsStock.ToDate2(dtpProdDate.Value);

			try
			{
				switch (Type)
				{
					case 1:
						if (PartSonBO.Instance.CheckExist("PartCode", sonPlanModel.PartCode) == true)
						{
							int result = (int)SonPlanBO.Instance.Insert(sonPlanModel);
						}
						else
						{
							MessageBox.Show("Mã linh kiện không tồn tại !!");
							return false;
						}
						break;
					case 2:
						SonPlanBO.Instance.Update(sonPlanModel);
						break;
				}
			}
			catch (Exception e)
			{
				return false;
			}


			return true;
		}

		private void frmAddEditPlan_Load(object sender, EventArgs e)
		{
			LoadDataToForm();
		}

		private void btnSaveClose_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (SaveData())
			{
				PartSonModel partSonModel = new PartSonModel();
				/* LoadDataToForm();*/
				txbPartCode.Text = "";
				txbOrderCode.Text = "";
				txbNote.Text = "";
				txbConfirmCode.Text = "";
				txbLotSize.Value = 0;
				txbNG.Value = 0;
				txbOP.Value = 0;
				txtCNT.Value = 0;
				txbQtyPlan.Value = 0;
				txbRealProdQty.Value = 0;
				txbSaleCode.Text = "";
				txbShipTo.Text = "";
				txbShipVia.Text = "";
				txbWorkerCode.Text = "";
				dtpDateExported.Value = TextUtilsStock.ToDate3(sonPlanModel.DateExported);
				dtpPrintedDate.Value = TextUtilsStock.ToDate3(sonPlanModel.PrintedDate);
				dtpProdDate.Value = TextUtilsStock.ToDate3(sonPlanModel.ProdDate);
				this.Text = "Thêm kế hoạch";
				Type = 1;
			}
		}

		private void frmAddEditPlan_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveClose_Click(null, null);
		}

		private void catVaThemOiToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void labelControl17_Click(object sender, EventArgs e)
		{

		}
	}
}
