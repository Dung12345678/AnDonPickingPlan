using BMS.Business;
using BMS.Model;
using BMS.Utils;
using DevExpress.Data.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmProductForceGun : _Forms
    {
        public frmProductForceGun()
        {
            InitializeComponent();
        }

        private void frmProductForceGun_Load(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }

        private void LoadInfoSearch()
        {
            DataTable dt = ProductCheckForceBO.Instance.LoadDataFromSP("spGetProductCheckForce", "ProductForceCheck", null, null);
            grdData.DataSource = dt;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmCheckForceDetail frm = new frmCheckForceDetail();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }
        private void btnCopy_Click(object sender, EventArgs e)
        {

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            int ID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (ID == 0) return;
            ProductCheckForceModel model = ProductCheckForceBO.Instance.FindByPK(ID) as ProductCheckForceModel;
            if (model == null) return;
            string ProductCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProductCode));
            if (string.IsNullOrEmpty(ProductCode.Trim())) return;
            Expression expression = new Expression("ProductCode", ProductCode);
            ArrayList arrProduct = ProductBO.Instance.FindByExpression(expression);
            if (arrProduct.Count == 0) return;
            ProductModel product = arrProduct[0] as ProductModel;
            frmCheckForceDetail frm = new frmCheckForceDetail();
            frm.model = model;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa dữ liệu?"), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    ProductCheckForceBO.Instance.Delete(strID);
                    grvData.DeleteRow(grvData.FocusedRowHandle);
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
                }
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (grvData.RowCount > 0)
                { FolderBrowserDialog od = new FolderBrowserDialog();
                    if (od.ShowDialog() == DialogResult.OK)
                    {
                        TextUtils.ExportExcel(grvData,od.SelectedPath,string.Format("Cai dat sung Form DATA_{0}",DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void btnFindAll_Click(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }        

        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadInfoSearch();
            }
        }

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0 && btnEdit.Enabled == true)
                btnEdit_Click(null, null);
        } 

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            frmImportCheckForceExcel frm = new frmImportCheckForceExcel();
            frm.ShowDialog();
            LoadInfoSearch();
        }
    }
}
