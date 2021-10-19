using IE;
using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmProductionPlanDetail : _Forms
	{

		private bool _isAdd;
		public ProductionPlanModel ProductionPlan = new ProductionPlanModel();

		public frmProductionPlanDetail()
		{
			InitializeComponent();
		}
		private void ProductionPlanDetail_Load(object sender, EventArgs e)
		{
			loadData();
		}
		// gán từ form frmProceductionPlan sang Detail
		// lấy dữ liệu từ form 1 sang form 2 
		//  public void funData(TextBox txtForm1) { lb.Text = txtForm1.Text; }
		// kiểm tra trước khi cất
		private bool ValidateForm()
		{
			if (txtProductcode.Text == "")
			{
				MessageBox.Show("Xin hãy điền mã ID", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			else
			{
				DataTable dt;
				if (ProductionPlan.ID > 0)
				{
					dt = LibIE.Select("select top 1 ProductionPlan from ProductionPlan where ProductCode = '" + txtProductcode.Text.Trim() + "' and ID <> " + ProductionPlan.ID);
				}
				else
				{
					dt = LibIE.Select("select top 1 ProductionPlan from ProductionPlan where ProductCode = '" + txtProductcode.Text.Trim() + "'");
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
			if (txtStt.Text == "")
			{
				MessageBox.Show("Xin hãy điền số thứ tự", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtAssembly.Text == "")
			{
				MessageBox.Show("Xin hãy điền Assembly order & cnt", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (txtOrderCode.Text == "")
			{
				MessageBox.Show("Xin hãy điền ASSY ORDER", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtCnt.Text == "")
			{
				MessageBox.Show("Xin hãy điền CNT", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtDescription.Text == "")
			{
				MessageBox.Show("Xin hãy điền Description", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtQty.Text == "")
			{
				MessageBox.Show("Xin hãy điền Qty", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtSalesOrder.Text == "")
			{
				MessageBox.Show("Xin hãy điền SalesOrder", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtShipTo.Text == "")
			{
				MessageBox.Show("Xin hãy điền ShipTo", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (txtSVia.Text == "")
			{
				MessageBox.Show("Xin hãy điền S.Via", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}

		bool saveData()
		{
			if (!ValidateForm()) return false;
			ProductionPlan.Stt = TextUtils.ToInt(txtStt.Text);
			ProductionPlan.AssemblyProduct = txtAssembly.Text.Trim();
			ProductionPlan.JgDate = TextUtils.ToDate3(dtpJgDate.Value);
			ProductionPlan.OrderCode = txtOrderCode.Text.Trim();
			ProductionPlan.Cnt = TextUtils.ToInt(txtCnt.Text);
			ProductionPlan.ProductCode = txtProductcode.Text.Trim();
			ProductionPlan.Description = txtDescription.Text.Trim();
			ProductionPlan.Qty = TextUtils.ToInt(txtQty.Text);
			ProductionPlan.SalesOrder = txtSalesOrder.Text.Trim();
			ProductionPlan.ShipTo = txtShipTo.Text.Trim();
			ProductionPlan.SVia = txtSVia.Text.Trim();
			ProductionPlan.Note = txtNote.Text.Trim();
			ProductionPlan.Code = txtCode.Text.Trim();
			ProductionPlan.OrderNew = txtOrderNew.Text.Trim();
			ProductionPlan.ProductCodeNew = txtProductCodeNew.Text.Trim();
			ProductionPlan.RequestDate = TextUtils.ToDate3(dtpRequestDate.Value);
			ProductionPlan.MaMoto = txtMaMoto.Text.Trim();
			ProductionPlan.MaMoto1 = txtMaMoto1.Text.Trim();
			ProductionPlan.DeliveryDate = TextUtils.ToDate3(dtpDeliveryDate.Value);
			ProductionPlan.AssemblyDate = TextUtils.ToDate3(dtpAssemblyDate.Value);
			ProductionPlan.OrderCodeFull = txtOrderCode.Text.Trim() + txtCnt.Text.Trim();
			ProductionPlan.Prints = checkPrints.Checked;

			if (ProductionPlan.ID > 0)
			{
				ProductionPlan.UpdatedAt = DateTime.Now;
				ProductionPlanBO.Instance.Update(ProductionPlan);
				MessageBox.Show("Update thành công").ToString();
			}
			else
			{
				ProductionPlan.CreatedAt = DateTime.Now;
				ProductionPlan.UpdatedAt = DateTime.Now;
				ProductionPlanBO.Instance.Insert(ProductionPlan);
				MessageBox.Show("Thêm mới thành công").ToString();
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
			if (saveData())
			{
				ProductionPlan = new ProductionPlanModel();
				loadData();
			}
		}
		void loadData()
		{
			txtStt.Text = TextUtils.ToString(ProductionPlan.Stt);
			txtProductcode.Text = ProductionPlan.ProductCode;
			txtAssembly.Text = ProductionPlan.AssemblyProduct;
			dtpJgDate.Value = TextUtils.ToDate3(ProductionPlan.JgDate);
			txtOrderCode.Text = ProductionPlan.OrderCode;
			txtCnt.Text = TextUtils.ToString(ProductionPlan.Cnt);
			txtDescription.Text = ProductionPlan.Description;
			txtQty.Text = TextUtils.ToString(ProductionPlan.Qty);
			txtSalesOrder.Text = ProductionPlan.SalesOrder;
			txtShipTo.Text = ProductionPlan.ShipTo;
			txtSVia.Text = ProductionPlan.SVia;
			txtNote.Text = ProductionPlan.Note;
			txtOrderNew.Text = ProductionPlan.OrderNew;
			txtProductCodeNew.Text = ProductionPlan.ProductCodeNew;
			txtCode.Text = ProductionPlan.Code;
			dtpRequestDate.Value = TextUtils.ToDate3(ProductionPlan.RequestDate); ;
			txtMaMoto.Text = ProductionPlan.MaMoto;
			txtMaMoto1.Text = ProductionPlan.MaMoto1;
			dtpDeliveryDate.Value = TextUtils.ToDate3(ProductionPlan.DeliveryDate);
			dtpAssemblyDate.Value = TextUtils.ToDate3(ProductionPlan.AssemblyDate);
			checkPrints.Checked = ProductionPlan.Prints;
		}

		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSave_Click(null, null);
		}

		private void saveNewToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveNew_Click(null, null);
		}

		private void lb_Click(object sender, EventArgs e)
		{
			//string sql = "select * from ProductionPlan where ID=" +lb.Text;
			//sql.
		}

		private void closeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}
