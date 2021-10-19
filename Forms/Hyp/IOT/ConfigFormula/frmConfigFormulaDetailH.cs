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
    public partial class ConfigFormulaDetailH : _Forms
    {
        public ConfigFormulaModel formulaModel = new ConfigFormulaModel();
        public ConfigFormulaDetailH()
        {
            InitializeComponent();
        }

        private void ConfigFormulaDetail_Load(object sender, EventArgs e)
        {
            loadcboProductGroup();
            loadConfigDetail();
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
            //cboProductTypeCode.ValueMember = "ID";
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
                formulaModel = new ConfigFormulaModel();
                loadConfigDetail();
            }
        }
        void loadConfigDetail()
        {
            cboProductGroupCode.EditValue = formulaModel.ProductGroupID;
            cboProductTypeCode.Text = formulaModel.ProductTypeCode;
            //  cboProductTypeCode.EditValue = formulaModel.ProductID;
            rtbFormula1.Text = formulaModel.Formula1;
            rtbFormula2.Text = formulaModel.Formula2;
            rtbFormula3.Text = formulaModel.Formula3;
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
            string sql1 = "";
            if (formulaModel.ID > 0)
            {
                sql1 = $"SELECT top 1 ID from ConfigFormula WHERE ProductTypeCode = N'{cboProductTypeCode.Text.Trim()}' and ProductGroupID={cboProductGroupCode.EditValue} and ID <>'{formulaModel.ID}'";
            }
            else
            {
                sql1 = $"SELECT top 1 ID from ConfigFormula WHERE ProductTypeCode = '{cboProductTypeCode.Text.Trim()}' and ProductGroupID={cboProductGroupCode.EditValue}";
            }
            int id = TextUtils.ToInt(LibIE.ExcuteScalar(sql1));
            if (id > 0)
            {
                MessageBox.Show("Đã tồn tại loại sản phẩm", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            return true;
        }
        bool saveData()
        {
            if (!ValidateForm()) return false;
            formulaModel.ProductGroupID = Lib.ToInt(cboProductGroupCode.EditValue);
            //formulaModel.ProductID = Lib.ToInt(cboProductTypeCode.EditValue);
            formulaModel.ProductGroupCode = cboProductGroupCode.Text.Trim();
            formulaModel.ProductTypeCode = cboProductTypeCode.Text.Trim();
            formulaModel.Formula1 = rtbFormula1.Text.Trim();
            formulaModel.Formula2 = rtbFormula2.Text.Trim();
            formulaModel.Formula3 = rtbFormula3.Text.Trim();
            if (formulaModel.ID > 0)
            {
                ConfigFormulaBO.Instance.Update(formulaModel);
            }
            else
            {
                formulaModel.ID = (int)ConfigFormulaBO.Instance.Insert(formulaModel);
            }
            return true;
        }


        private void ConfigFormulaDetail_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
        private void cboProductGroupCode_TextChanged(object sender, EventArgs e)
        {
            loadcboProductType();
        }

		private void saveToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            btnSave_Click(null, null);
		}

		private void saveNewToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            btnSaveNew_Click(null, null);
		}

		private void closeToolStripMenuItem_Click_1(object sender, EventArgs e)
		{
            this.DialogResult = DialogResult.OK;
        }
	}
}
