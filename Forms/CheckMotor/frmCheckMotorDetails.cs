using SD.Business;
using SD.Model;
using SD.Facade;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SD.Utils;
using System.Collections;

namespace BMS
{
	public partial class frmCheckMotorDetails : _Forms
	{
		public CheckMotorModel _CheckMotorModel = new CheckMotorModel();
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

		public frmCheckMotorDetails(int typeEvent)
		{
			InitializeComponent();
			Type = typeEvent;
			switch (typeEvent)
			{
				case 1:
					this.Text = "Thêm Motor";
					break;
				case 2:
					this.Text = "Sửa Motor";
					break;
			}
		}
		private bool ValidateForm()
		{
			if (txtCardNo.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy nhập Card No", TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;

		}
		void LoadData()
		{
			txtCardNo.Text = _CheckMotorModel.CardNo;
			txtArticleID.Text = _CheckMotorModel.ArticleID;
			txtSalesOrder.Text = _CheckMotorModel.SalesOrder;
			txtDescription.Text = _CheckMotorModel.Descriptions;
			txtMotorInspSealNo.Text = _CheckMotorModel.MotorInspSealNo;
			txtSerialNo.Text = _CheckMotorModel.SerialNo;
			txtOrderedQty.Text = TextUtils.ToString(_CheckMotorModel.OrderedQty);
			txtAssemblyOrderNo.Text = _CheckMotorModel.AssemblyOrderNo;
			txtUPR.Text = _CheckMotorModel.UPR;
			txtQtyOfShipOrder.Text = TextUtils.ToString(_CheckMotorModel.QtyOfShipOrder);
			txtNoOfJG.Text = TextUtils.ToString(_CheckMotorModel.NoOfJG);
		}
		bool SaveData()
		{

			if (!ValidateForm()) return false;
			//_CheckMotorModel.CardNo = txtCardNo.Text.Trim();
			_CheckMotorModel.ArticleID = txtArticleID.Text.Trim();
			_CheckMotorModel.SalesOrder = txtSalesOrder.Text.Trim();
			_CheckMotorModel.Descriptions = txtDescription.Text.Trim();
			_CheckMotorModel.MotorInspSealNo = txtMotorInspSealNo.Text.Trim();
			_CheckMotorModel.SerialNo = txtSerialNo.Text.Trim();
			_CheckMotorModel.OrderedQty = TextUtils.ToInt(txtOrderedQty.Text.Trim());
			_CheckMotorModel.AssemblyOrderNo = txtAssemblyOrderNo.Text.Trim();
			_CheckMotorModel.UPR = txtUPR.Text.Trim();
			_CheckMotorModel.QtyOfShipOrder = TextUtils.ToInt(txtQtyOfShipOrder.Text.Trim());
			_CheckMotorModel.NoOfJG = TextUtils.ToInt(txtNoOfJG.Text.Trim());
			Expression exp1 = new Expression("CardNo", _CheckMotorModel.CardNo);
			Expression exp2 = new Expression("MotorInspSealNo", _CheckMotorModel.MotorInspSealNo);
			Expression exp3 = new Expression("SerialNo", _CheckMotorModel.SerialNo);

			ArrayList arr = CheckMotorBO.Instance.FindByExpression(exp1.And(exp2).And(exp3));
			try
			{
				if (arr.Count > 0)
				{
					_CheckMotorModel.JGDate = DateTime.Now;
					CheckMotorBO.Instance.Update(_CheckMotorModel);
					MessageBox.Show("Update thành công").ToString();
				}
				else
				{
					_CheckMotorModel.JGDate = DateTime.Now;
					CheckMotorBO.Instance.Insert(_CheckMotorModel);
					MessageBox.Show("Thêm mới thành công").ToString();
				}
			}
			catch
			{

			}
			return true;
		}

		private void frmAddEditPlan_Load(object sender, EventArgs e)
		{
			LoadData();
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
				_CheckMotorModel = new CheckMotorModel();
				LoadData();
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
	}
}
