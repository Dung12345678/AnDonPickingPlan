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
    public partial class frmUserGroup : _Forms
    {
        public frmUserGroup()
        {
            InitializeComponent();
        }

        private void frmUserGroup_Load(object sender, EventArgs e)
        {
            loadData();
        }

        private void loadData()
        {
            DataTable dt = TextUtils.Select("SELECT * FROM dbo.UserGroup WITH(NOLOCK)");
            grdData.DataSource = dt;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmUserGroupDetail frm = new frmUserGroupDetail();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                loadData();
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            try
            {
                int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
                if (id == 0) return;
                UserGroupModel model = (UserGroupModel)UserGroupBO.Instance.FindByPK(id);
                frmUserGroupDetail frm = new frmUserGroupDetail();
                frm.UserGroup = model;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    loadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (grvData.FocusedRowHandle < 0)
                return;

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));

            string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("Name"));

            if (UserGroupLinkBO.Instance.CheckExist("UserGroupID", strID))
            {
                MessageBox.Show("Nhóm quyền này đã phát sinh ở các nghiệp vụ khác nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (UserGroupRightDistributionBO.Instance.CheckExist("UserGroupID", strID))
            {
                MessageBox.Show("Nhóm quyền này đã phát sinh ở các nghiệp vụ khác nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    UserGroupBO.Instance.Delete(strID);
                    grvData.DeleteSelectedRows();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnRole_Click(object sender, EventArgs e)
        {
            frmUserGroupPermission frm = new frmUserGroupPermission();
            frm.GroupID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));
            frm.Show();
        }

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            btnEdit_Click(null, null);
        }
        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadDetail();
        }

        void loadDetail()
        {
            int masterID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));
            if (masterID == 0) return;

            DataTable dt = TextUtils.GetDataTableFromSP("spGetUserGroupLink", new string[] { "@UserID", "@UserGroupID" }, new object[] { 0, masterID });
            grdDetail.DataSource = dt;
        }

        private void btnChooseUser_Click(object sender, EventArgs e)
        {
            frmChooseEmployee frm = new frmChooseEmployee();
            if (frm.ShowDialog()== DialogResult.OK)
            {
                List<int> lstEmployeeID = frm.LstID;
                if (lstEmployeeID.Count == 0) return;
                int masterID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));
                if (masterID == 0) return;

                string lstID = string.Join(";", lstEmployeeID);

                TextUtils.ExcuteProcedure("spAddUserToGroupRole", new string[] { "@ListUserID", "@UserGroupID" }, new object[] { lstID, masterID });
                loadDetail();
            }
        }

        private void btnDeleteUser_Click(object sender, EventArgs e)
        {
            if (grvDetail.FocusedRowHandle < 0)
                return;

            int strID = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue("ID"));

            string strName = TextUtils.ToString(grvDetail.GetFocusedRowCellValue("FullName"));
           
            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    UserGroupLinkBO.Instance.Delete(strID);
                    grvDetail.DeleteSelectedRows();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
       
    }
}
