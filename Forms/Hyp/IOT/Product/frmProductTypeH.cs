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
    public partial class frmProductTypeH : _Forms
    {
        private bool _isAdd;
        public frmProductTypeH()
        {
            InitializeComponent();
        }

        private void frmProductTypeH_Load(object sender, EventArgs e)
        {
            loadData();
        }

        #region Methods
        private void setInterface(bool isEdit)
        {
            txtCode.ReadOnly = !isEdit;

            grdData.Enabled = !isEdit;

            btnSave.Visible = isEdit;
            btnCancel.Visible = isEdit;

            btnNew.Visible = !isEdit;
            btnEdit.Visible = !isEdit;
            btnDelete.Visible = !isEdit;
        }

        private void clearInterface()
        {
            txtName.Text = "";
            txtCode.Text = "";
           
        }

        private bool checkValid()
        {
            if (txtCode.Text == "")
            {
                MessageBox.Show("Xin hãy điền mã của loại sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (!_isAdd)
                {
                    int strID = TextUtils.ToInt(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID"));
                    dt = LibIE.Select("select top 1 1 from ProductType WITH(NOLOCK) where ProductTypeCode = '" + txtCode.Text.Trim() + "' and ID <> " + strID);
                }
                else
                {
                    dt = LibIE.Select("select top 1 1 from ProductType WITH(NOLOCK) where ProductTypeCode = '" + txtCode.Text.Trim() + "'");
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
            if (txtName.Text == "")
            {
                MessageBox.Show("Xin hãy điền tên của loại sản phẩm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            return true;
        }

        private void loadData()
        {
            try
            {
                DataTable tbl = LibIE.Select("SELECT * FROM ProductType WITH(NOLOCK)");
                grdData.DataSource = tbl;
            }
            catch (Exception)
            {
            }
        }
        
        #endregion

        #region Buttons Events
        private void btnNew_Click(object sender, EventArgs e)
        {
            setInterface(true);
            _isAdd = true;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;
            setInterface(true);
            _isAdd = false;
            txtName.Text = TextUtils.ToString(grvData.GetFocusedRowCellValue("ProductTypeName"));
            txtCode.Text = TextUtils.ToString(grvData.GetFocusedRowCellValue("ProductTypeCode"));
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));

            string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("ProductTypeName"));

            if (ProductBO.Instance.CheckExist("ProductTypeID", strID))
            {
                MessageBox.Show("Loại sản phẩm này đang được sử dụng bên sản phẩm nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    ProductTypeBO.Instance.Delete(Convert.ToInt32(strID));
                    loadData();
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!checkValid()) return;

                ProductTypeModel dModel;
                if (_isAdd)
                {
                    dModel = new ProductTypeModel();
                }
                else
                {
                    int ID = Convert.ToInt32(grvData.GetFocusedRowCellValue("ID"));
                    dModel = (ProductTypeModel)ProductTypeBO.Instance.FindByPK(ID);
                }

                dModel.ProductTypeCode = txtCode.Text.Trim();
                dModel.ProductTypeName = txtName.Text.Trim();

                if (_isAdd)
                {
                    ProductTypeBO.Instance.Insert(dModel);
                }
                else
                    ProductTypeBO.Instance.Update(dModel);

                loadData();
                setInterface(false);
                clearInterface();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            setInterface(false);
            clearInterface();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private void grdData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0 && btnEdit.Enabled == true)
                btnEdit_Click(null, null);
        }
    }
}
