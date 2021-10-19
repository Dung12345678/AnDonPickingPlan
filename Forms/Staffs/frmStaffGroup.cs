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
    public partial class frmStaffGroup : _Forms
    {
        private bool isAdd;
        public frmStaffGroup()
        {
            InitializeComponent();
        }

        private void frmStaffGroup_Load(object sender, EventArgs e)
        {
            try
            {
                loadCombo();
                LoadData();
            }
            catch (Exception ex)
            {
                TextUtils.ShowError(ex);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Methods
        private void SetInterface(bool isEdit)
        {
            txtCode.ReadOnly = !isEdit;

            grdData.Enabled = !isEdit;

            btnSave.Visible = isEdit;
            btnCancel.Visible = isEdit;

            btnNew.Visible = !isEdit;
            btnEdit.Visible = !isEdit;
            btnDelete.Visible = !isEdit;
        }

        private void ClearInterface()
        {
            txtName.Text = "";
            txtCode.Text = "";
            cboDepartment.Text = "";
        }

        private bool checkValid()
        {
            if (txtCode.Text == "")
            {
                MessageBox.Show("Xin hãy điền mã nhóm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (!isAdd)
                {
                    int strID = TextUtils.ToInt(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());
                    dt = TextUtils.Select("select Code from UserGroup where Code = '" + txtCode.Text.Trim() + "' and ID <> " + strID);
                }
                else
                {
                    dt = TextUtils.Select("select Code from UserGroup where Code = '" + txtCode.Text.Trim() + "'");
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
                MessageBox.Show("Xin hãy điền tên nhóm.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            //if (cboDepartment.EditValue == null)
            //{
            //    MessageBox.Show("Xin hãy chọn phòng ban.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return false;
            //}

            return true;
        }

        void loadCombo()
        {
            try
            {
                DataTable tblDepartment = TextUtils.Select("Select ID, Code, Name from Department a with(nolock)");
                TextUtils.PopulateCombo(cboDepartment, tblDepartment.Copy(), "Name", "ID", "");
            }
            catch
            {
            }
        }

        private void LoadData()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select * from UserGroup with(nolock)");
                grdData.DataSource = tbl;                
                grvData.BestFitColumns();

                DataTable tblPerson = TextUtils.Select("Select ID, Code, Name from Department a with(nolock)");
                repositoryItemLookUpEdit1.DataSource = tblPerson;
                repositoryItemLookUpEdit1.DisplayMember = "Name";
                repositoryItemLookUpEdit1.ValueMember = "ID";
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Buttons Events
        private void btnNew_Click(object sender, EventArgs e)
        {
            SetInterface(true);
            isAdd = true;
            //cboStatus.SelectedIndex = 1;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;
            SetInterface(true);
            isAdd = false;

            int ID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            UserGroupModel dModel = (UserGroupModel)UserGroupBO.Instance.FindByPK(ID);

            txtName.Text = dModel.Name;
            txtCode.Text = dModel.Code;
            cboDepartment.SelectedValue = dModel.DepartmentID;
            //cboStatus.SelectedIndex = dModel.Status;
            //leHead.EditValue = dModel.UserID;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());

            string strName = grvData.GetRowCellValue(grvData.FocusedRowHandle, "Name").ToString();

            //if (WorksBO.Instance.CheckExist("CustomerID", strID))
            //{
            //    MessageBox.Show("Khách hàng này đang được sử dụng nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return;
            //}

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    UserGroupBO.Instance.Delete(Convert.ToInt32(strID));
                    LoadData();
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
                if (checkValid())
                {
                    UserGroupModel dModel;
                    if (isAdd)
                    {
                        dModel = new UserGroupModel();
                    }
                    else
                    {
                        int ID = Convert.ToInt32(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());
                        dModel = (UserGroupModel)UserGroupBO.Instance.FindByPK(ID);
                    }
                    dModel.Code = txtCode.Text;
                    dModel.Name = txtName.Text;
                    dModel.DepartmentID = TextUtils.ToInt(cboDepartment.SelectedValue);
                    if (isAdd)
                    {
                        UserGroupBO.Instance.Insert(dModel);
                    }
                    else
                        UserGroupBO.Instance.Update(dModel);

                    LoadData();
                    SetInterface(false);
                    ClearInterface();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            SetInterface(false);
            ClearInterface();
        }

        #endregion

        private void grdData_Click(object sender, EventArgs e)
        {

        }

        private void grdData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0 && btnEdit.Enabled == true)
                btnEdit_Click(null, null);
        }

    }
}
