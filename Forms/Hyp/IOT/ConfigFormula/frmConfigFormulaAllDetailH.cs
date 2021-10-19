using IE.Business;
using IE.Model;
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
	public partial class frmConfigFormulaAllDetailH : _Forms
	{
		public ConfigFormulaAllModel formulaModel = new ConfigFormulaAllModel();
		public frmConfigFormulaAllDetailH()
		{
			InitializeComponent();
		}

		private void frmConfigFormulaAllDetailH_Load(object sender, EventArgs e)
		{
			loadcboProductGroup();
			loadConfigDetail();
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnSaveNew_Click(object sender, EventArgs e)
		{
			if (saveData())
			{
				formulaModel = new ConfigFormulaAllModel();
				loadConfigDetail();
			}
		}
		void loadcboProductGroup()
		{
			DataTable dt = new DataTable();
			dt = LibIE.Select("SELECT * FROM ProductGroup");
			cboProductGroupCode.Properties.DisplayMember = "ProductGroupName";
			cboProductGroupCode.Properties.ValueMember = "ID";
			cboProductGroupCode.Properties.DataSource = dt;
		}
		void loadcboProductType()
		{
			DataTable dt = new DataTable();
			dt = LibIE.Select(string.Format($"SELECT DISTINCT ProductTypeCode  FROM dbo.Product where ProductGroupID='{cboProductGroupCode.EditValue}'order by ProductTypeCode"));
			cboProductTypeCode.DataSource = dt;
			cboProductTypeCode.DisplayMember = "ProductTypeCode";
			cboProductTypeCode.Text = "";
		}
		void loadConfigDetail()
		{
			cboProductGroupCode.EditValue = formulaModel.ProductGroupID;
			cboProductTypeCode.Text = formulaModel.ProductTypeCode;
			txtSTT.Text = Lib.ToString(formulaModel.STT);
			rtbFormula.Text = formulaModel.Formula;
		}
		private bool ValidateForm()
		{
			if (cboProductGroupCode.EditValue == null)
			{
				MessageBox.Show("Xin hãy chọn Group.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (cboProductTypeCode.Text.Trim() == "")
			{
				MessageBox.Show("Xin hãy chọn loại sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			return true;
		}
		bool saveData()
		{
			if (!ValidateForm()) return false;
			formulaModel.ProductGroupID = Lib.ToInt(cboProductGroupCode.EditValue);
			formulaModel.ProductGroupCode = cboProductGroupCode.Text.Trim();
			formulaModel.STT = Lib.ToInt(txtSTT.Text.Trim());
			formulaModel.ProductTypeCode = cboProductTypeCode.Text.Trim();
			formulaModel.Formula = rtbFormula.Text.Trim();

			if (formulaModel.ID > 0)
			{
				ConfigFormulaAllBO.Instance.Update(formulaModel);
			}
			else
			{
				formulaModel.ID = (int)ConfigFormulaAllBO.Instance.Insert(formulaModel);
			}
			return true;
		}

		private void cboProductGroupCode_TextChanged(object sender, EventArgs e)
		{
			loadcboProductType();
		}
	}
}
