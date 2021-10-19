using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Utils;
using BMS.Model;
using BMS.Business;
using BMS;

namespace BMS
{
    public partial class frmDepartment : _Forms
    {
        private bool isAdd;

        public frmDepartment()
        {
            InitializeComponent();
        }

        private void frmDepartment_Load(object sender, EventArgs e)
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

        #region Methods
        private void SetInterface(bool isEdit)
        {
            txtCode.ReadOnly = !isEdit;
            cboStatus.Enabled = isEdit;

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
            cboStatus.SelectedIndex = 1;
            txtCode.Text = "";
            txtEmail.Text = "";
            leHead.EditValue = 0;
        }

        private bool checkValid()
        {
            if (txtCode.Text == "")
            {
                MessageBox.Show("Xin hãy điền mã của phòng ban.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (!isAdd)
                {
                    int strID = TextUtils.ToInt(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());
                    dt = TextUtils.Select("select Code from Department where Code = '" + txtCode.Text.Trim() + "' and ID <> " + strID );                    
                }
                else
                {
                    dt = TextUtils.Select("select Code from Department where Code = '" + txtCode.Text.Trim() + "'");                    
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
                MessageBox.Show("Xin hãy điền tên của phòng ban.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            if (cboStatus.SelectedIndex < 0)
            {
                MessageBox.Show("Xin hãy chọn trạng thái của phòng ban.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
           
            return true;
        }

        void loadCombo()
        {
            try
            {
                //load lookupedit
                DataTable tblPerson = TextUtils.Select("Select ID, Code, FullName from Users a with(nolock)");
                leHead.Properties.DataSource = tblPerson;
                leHead.Properties.DisplayMember = "FullName";
                leHead.Properties.ValueMember = "ID";
                //TextUtils.PopulateCombo(leHead, tblPerson.Copy(), "Name", "ID", "-- Chọn trưởng phòng --");
            }
            catch (Exception)
            {
            }           
        }
        private void LoadData()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select a.*, Case when a.Status = 0 then N'Ngừng hoạt động' else N'Hoạt động' end StatusText from Department a with(nolock)");
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
            SetInterface(true);
            isAdd = true;
            cboStatus.SelectedIndex = 1;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;
            SetInterface(true);
            isAdd = false;
            txtName.Text = grvData.GetRowCellValue(grvData.FocusedRowHandle, "Name").ToString();
            txtCode.Text = grvData.GetRowCellValue(grvData.FocusedRowHandle, "Code").ToString();
            cboStatus.SelectedIndex = Convert.ToInt16(grvData.GetRowCellValue(grvData.FocusedRowHandle, "Status"));
            txtEmail.Text = grvData.GetFocusedRowCellValue(colEmail).ToString();
            leHead.EditValue = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colHead));
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());

            string strName = grvData.GetRowCellValue(grvData.FocusedRowHandle, "Name").ToString();

            if (UsersBO.Instance.CheckExist("DepartmentID", strID))
            {
                MessageBox.Show("Phòng ban này đang được sử dụng nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {                
                    DepartmentBO.Instance.Delete(Convert.ToInt32(strID));
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
                    DepartmentModel dModel;
                    if (isAdd)
                    {
                        dModel = new DepartmentModel();                        
                        dModel.CreatedDate = TextUtils.GetSystemDate();
                        dModel.UpdatedDate = dModel.CreatedDate;
                        dModel.CreatedBy = Global.AppUserName;
                        dModel.UpdatedBy = Global.AppUserName;                       
                    }
                    else
                    {
                        int ID = Convert.ToInt32(grvData.GetRowCellValue(grvData.FocusedRowHandle, "ID").ToString());
                        dModel = (DepartmentModel)DepartmentBO.Instance.FindByPK(ID);                       
                        dModel.UpdatedDate = TextUtils.GetSystemDate();
                        dModel.UpdatedBy = Global.AppUserName;                        
                    }
                    dModel.Code = txtCode.Text;
                    dModel.Name = txtName.Text;
                    dModel.Status = cboStatus.SelectedIndex;
                    dModel.Email = txtEmail.Text;
                    dModel.HeadofDepartment = TextUtils.ToInt(leHead.EditValue);

                    if (isAdd)
                    {
                        DepartmentBO.Instance.Insert(dModel);
                    }
                    else
                        DepartmentBO.Instance.Update(dModel);

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

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private void grdData_Click(object sender, EventArgs e)
        {
            try
            {
                if (grvData.RowCount > 0)
                {
                    lblCreatedBy.Text = grvData.GetRowCellValue(grvData.FocusedRowHandle, "CreatedBy").ToString();
                    lblCreatedDate.Text = ((DateTime)grvData.GetRowCellValue(grvData.FocusedRowHandle, "CreatedDate")).ToString(TextUtils.FomatShortDate);
                    lblUpdatedBy.Text = grvData.GetRowCellValue(grvData.FocusedRowHandle, "UpdatedBy").ToString();
                    lblUpdatedDate.Text = ((DateTime)grvData.GetRowCellValue(grvData.FocusedRowHandle, "UpdatedDate")).ToString(TextUtils.FomatShortDate);
                }
                else
                {
                    lblCreatedBy.Text = "";
                    lblCreatedDate.Text = "";
                    lblUpdatedBy.Text = "";
                    lblUpdatedDate.Text = "";
                }
            }
            catch
            {
                lblCreatedBy.Text = "";
                lblCreatedDate.Text = "";
                lblUpdatedBy.Text = "";
                lblUpdatedDate.Text = "";
            }
        }

        private void grdData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0 && btnEdit.Enabled == true)
                btnEdit_Click(null, null);
        }
    }
}
