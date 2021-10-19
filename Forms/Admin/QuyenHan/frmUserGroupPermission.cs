using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Business;
using BMS.Utils;
using System.Reflection;
using System.Threading;
using BMS.Model;

namespace BMS
{
    public partial class frmUserGroupPermission : _Forms
    {
        public int GroupID = 0;
        public frmUserGroupPermission()
        {
            InitializeComponent();
        }

        private void frmUserGroupPermission_Load(object sender, EventArgs e)
        {
            loadStaffGroup();
            loadTree();
            //treeData.Focus();
        }

        void loadStaffGroup()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select * from UserGroup a with(nolock)");
                cboUserGroup.DataSource = tbl;
                cboUserGroup.DisplayMember = "Name";
                cboUserGroup.ValueMember = "ID";

                if (GroupID > 0)
                {
                    cboUserGroup.SelectedValue = GroupID;
                }
            }
            catch
            {
            }
        }

        void loadTree()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select ID,Name,ParentID from FormAndFunctionGroup with(nolock) order by Name");

                DataRow row = tbl.NewRow();
                row["ID"] = 0;
                row["Name"] = "--Tất cả các nhóm--";
                tbl.Rows.InsertAt(row, 0);

                treeData.DataSource = tbl;
                treeData.KeyFieldName = "ID";
                treeData.PreviewFieldName = "Name";
                treeData.ParentFieldName = "ParentID";

                treeData.ExpandAll();
                //treeData.CollapseAll();
            }
            catch
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        void loadPermission()
        {
            try
            {
                if (cboUserGroup.SelectedIndex < 0) return;
                int userGroupID = TextUtils.ToInt(cboUserGroup.SelectedValue);
                int catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
                if (catID == 0) return;

                DataTable dt = TextUtils.GetDataTableFromSP("spGetUserGroupPermission", 
                    new string[] {"@UserGroupID", "@FormAndFunctionGroupID" }, 
                    new object[] { userGroupID, catID });
                grdData.DataSource = dt;
            }
            catch
            {
            }
        }
        private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
        {
            loadPermission();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (cboUserGroup.SelectedIndex < 0) return;
            int userGroupID = TextUtils.ToInt(cboUserGroup.SelectedValue);
            int catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
            if (catID == 0) return;
            grvData.FocusedRowHandle = -1;

            List<int> lstPerID = new List<int>();
            for (int i = 0; i < grvData.RowCount; i++)
            {
                bool isRole = TextUtils.ToBoolean(grvData.GetRowCellValue(i, colIsHasRole));
                if (isRole)
                {
                    lstPerID.Add(TextUtils.ToInt(grvData.GetRowCellValue(i, colID)));
                }
            }

            TextUtils.ExcuteProcedure("spAddPermissionToGroupRole",
                new string[] { "@UserGroupID", "@FormAndFunctionGroupID", "@ListPermissionID" },
                new object[] { userGroupID, catID, lstPerID.Count > 0 ? string.Join(";", lstPerID) : "" });

            MessageBox.Show("Phân quyền thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void cboUserGroup_SelectionChangeCommitted(object sender, EventArgs e)
        {
            loadPermission();
        }
    }
}
