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
    public partial class frmPermissionGroup : _Forms
    {
        public frmPermissionGroup()
        {
            InitializeComponent();
        }

        private void frmPermissionGroup_Load(object sender, EventArgs e)
        {
            loadStaffGroup();
            loadTree();
        }

        void loadStaffGroup()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select * from vUserGroup a with(nolock)");
                grdData.DataSource = tbl;
                grvData.BestFitColumns();
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
                int _catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
                if (_catID == 0) return;
                if (TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID)) == 0) return;

                DataTable Source = TextUtils.Select("Select * from FormAndFunction WITH(NOLOCK) where FormAndFunctionGroupID = " + _catID);
                //DataTable Source = TextUtils.Select("Select * from vFormAndFunction WITH(NOLOCK)");
                treePermission.DataSource = Source;
                treePermission.KeyFieldName = "ID";
                treePermission.PreviewFieldName = "Name";
                treePermission.ParentFieldName = "FormAndFunctionGroupID";

                for (int i = 0; i < treePermission.Nodes.Count; i++)
                {
                    DataTable dt = TextUtils.Select("Select * from UserGroupRightDistribution WITH(NOLOCK) where FormAndFunctionID = "
                        + TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer)) + " and UserGroupID = "
                        + TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID)));
                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            treePermission.GetNodeByVisibleIndex(i).Checked = true;
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadPermission();
        }

        private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
        {
            loadPermission();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (treePermission.Nodes.Count == 0) return;
            int userGroupID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (userGroupID == 0) return;

            for (int i = 0; i < treePermission.Nodes.Count; i++)
            {
                int perID = TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer));

                DataTable dt = TextUtils.Select("Select * from UserGroupRightDistribution WITH(NOLOCK) where FormAndFunctionID = "
                    + perID + " and UserGroupID = " + userGroupID);
                if (dt.Rows.Count > 0)
                {
                    if (!treePermission.GetNodeByVisibleIndex(i).Checked)
                    {
                        UserGroupRightDistributionBO.Instance.Delete(TextUtils.ToInt(dt.Rows[0]["ID"]));
                    }
                }
                else
                {
                    if (treePermission.GetNodeByVisibleIndex(i).Checked)
                    {
                        UserGroupRightDistributionModel model = new UserGroupRightDistributionModel();
                        model.FormAndFunctionID = TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer));
                        model.UserGroupID = userGroupID;
                        UserGroupRightDistributionBO.Instance.Insert(model);
                    }
                }
            }
            MessageBox.Show("Phân quyền thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
