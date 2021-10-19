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
    public partial class frmRole : _Forms
    {
        int _catID = 0;

        public frmRole()
        {
            InitializeComponent();
            LoadDepartment();
            loadStaff();
            loadTree();
        }

        private void LoadDepartment()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select a.* from Department a with(nolock) where a.Status = 1 ");
                leDepartment.Properties.DataSource = tbl;
                leDepartment.Properties.DisplayMember = "Name";
                leDepartment.Properties.ValueMember = "ID";
            }
            catch (Exception)
            {
            }

        }

        void loadStaff()
        {
            try
            {
                DataTable tbl = TextUtils.Select("Select a.* from Users a with(nolock) where Status = 0 and a.DepartmentID = " + TextUtils.ToInt(leDepartment.EditValue));
                grdData.DataSource = tbl;
                grvData.BestFitColumns();
            }
            catch (Exception)
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

                //treeData.ExpandAll();
                treeData.CollapseAll();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        void loadPermission()
        {
            try
            {
                _catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
                if (_catID == 0) return;
                if (TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID)) == 0) return;

                DataTable Source = TextUtils.Select("Select * from FormAndFunction WITH(NOLOCK) where FormAndFunctionGroupID = " + _catID );
                //DataTable Source = TextUtils.Select("Select * from vFormAndFunction WITH(NOLOCK)");
                treePermission.DataSource = Source;
                treePermission.KeyFieldName = "ID";
                treePermission.PreviewFieldName = "Name";
                treePermission.ParentFieldName = "FormAndFunctionGroupID";

                for (int i = 0; i < treePermission.Nodes.Count; i++)
                {
                    DataTable dt = TextUtils.Select("Select * from UserRightDistribution WITH(NOLOCK) where FormAndFunctionID = "
                        + TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer)) + " and UserID = "
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
            catch (Exception)
            {

            }
        }

        private void leDepartment_EditValueChanged(object sender, EventArgs e)
        {
            loadStaff();
        }

        private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
        {
            loadPermission();
        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadPermission();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (treePermission.Nodes.Count == 0) return;
            int userID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (userID == 0) return;

            for (int i = 0; i < treePermission.Nodes.Count; i++)
            {
                int perID = TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer));
              
                DataTable dt = TextUtils.Select("Select * from UserRightDistribution WITH(NOLOCK) where FormAndFunctionID = "
                    + perID + " and UserID = " + userID);
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        if (!treePermission.GetNodeByVisibleIndex(i).Checked)
                        {
                            UserRightDistributionBO.Instance.Delete(TextUtils.ToInt(dt.Rows[0][0]));
                        }
                    }
                    else
                    {
                        if (treePermission.GetNodeByVisibleIndex(i).Checked)
                        {
                            UserRightDistributionModel model = new UserRightDistributionModel();
                            model.FormAndFunctionID = TextUtils.ToInt(treePermission.GetNodeByVisibleIndex(i).GetValue(colIDPer));
                            model.UserID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
                            UserRightDistributionBO.Instance.Insert(model);
                        }
                    }
                }
            }
            MessageBox.Show("Phân quyền thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
