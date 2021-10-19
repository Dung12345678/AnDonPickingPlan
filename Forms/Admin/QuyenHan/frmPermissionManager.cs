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
using System.Collections;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System.Reflection;
using System.Threading;

namespace BMS
{
    public partial class frmPermissionManager : _Forms
    {
        int _curentNode = 0;
        frmCshLoading frmLoad = null;
        Thread thSearch = null;
        int _RownIndex = 0;
        int _catID = 0;

        public frmPermissionManager()
        {
            InitializeComponent();
            loadTree();
        }

        #region Methods
        bool isParent(int id)
        {
            try
            {
                DataTable tbl = TextUtils.Select(@"select * from dbo.FormAndFunctionGroup a WITH(NOLOCK) where ParentID = " + id + " ORDER BY Name");
                if (tbl.Rows.Count > 0)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
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

                //DevExpress.XtraTreeList.Nodes.TreeListNode tNode = treeData.FindNodeByFieldValue("ID", 0);
                //tNode.Expanded = true;

                //DataTable tblRoot = TextUtils.Select("Select a.ID from FormAndFunctionGroup a with(nolock) where ParentID = 0");
                //foreach (DataRow item in tblRoot.Rows)
                //{
                //    DevExpress.XtraTreeList.Nodes.TreeListNode t = treeData.FindNodeByFieldValue("ID", TextUtils.ToInt(item["ID"]));
                //    t.Expanded = true;
                //}

                DevExpress.XtraTreeList.Nodes.TreeListNode currentNode = treeData.FindNodeByFieldValue("ID", _curentNode);
                treeData.SetFocusedNode(currentNode);
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        private void LoadInfoSearch()
        {
            #region Parameter
            try
            {
                _catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            #endregion

            try
            {
                // Gọi Thread để xử lí dữ liệu
                thSearch = new Thread(new ThreadStart(ProcessLoadData));
                thSearch.Start();
                treeData.Select();
            }
            catch (Exception ex)
            {
                return;
            }
        }

        public void ProcessLoadData()
        {
            try
            {
                //Load trạng thái khi bắt đầu search.
                SetMode(true);
                //Truy xuất dữ liệu.                
                DataTable Source = TextUtils.Select("Select * from FormAndFunction WITH(NOLOCK) where FormAndFunctionGroupID = " + _catID + " ORDER BY Name");
                //Đổ dữ liệu lên grid              
                this.BeginInvoke((MethodInvoker)delegate
                {
                    grdData.DataSource = Source;
                    //Focus con trỏ chuột tại dòng đã select khi load lại dữ liệu
                    if (_RownIndex >= grvData.RowCount)
                        _RownIndex = 0;
                    if (_RownIndex > 0)
                        grvData.FocusedRowHandle = _RownIndex;
                    grvData.BestFitColumns();
                });
                //Kết thúc khi load xong dữ liệu
                SetMode(false);
                //Focus
                this.BeginInvoke((MethodInvoker)delegate { this.SetFocus(); });
            }
            catch (Exception ex)
            {
                SetMode(false);
                thSearch.Abort();
                return;
            }
        }

        private void SetMode(bool IsBeginLoad)
        {
            if (IsBeginLoad == true)
            {
                this.BeginInvoke((MethodInvoker)delegate { frmLoad = new frmCshLoading(); frmLoad.ShowDialog(); });
            }
            else
            {
                this.BeginInvoke((MethodInvoker)delegate { frmLoad.CloseForm(); });
            }

            //SetControlPropertyThreadSafe((Control)toolStrip1, "Enabled", !IsBeginLoad);
            //SetControlPropertyThreadSafe((Control)grdGrid, "Enabled", !IsBeginLoad);
            //SetControlPropertyThreadSafe((Control)groupControl1, "Enabled", !IsBeginLoad);
            //SetControlPropertyThreadSafe((Control)groupBox1, "Enabled", !IsBeginLoad);
            //SetControlPropertyThreadSafe((Control)groupBox2, "Enabled", !IsBeginLoad);
        }

        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);
        public static void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            }
            else
            {
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
            }
        }

        public void SetFocus()
        {
            //txtCode.Focus();
            treeData.Focus();
        }

        #endregion

        #region Events

        #region Permission Category Events
        private void btnNewGroup_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
            frmPermissionCategory frm = new frmPermissionCategory();
            frm.ParentID = id;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                _curentNode = frm.CurentNode;
                loadTree();
            }
        }

        private void btnEditGroup_Click(object sender, EventArgs e)
        {
            try
            {
                int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
                FormAndFunctionGroupModel model = (FormAndFunctionGroupModel)FormAndFunctionGroupBO.Instance.FindByPK(id);
                frmPermissionCategory frm = new frmPermissionCategory();
                frm.ParentID = model.ParentID;
                frm.Model = model;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    _curentNode = frm.CurentNode;
                    loadTree();
                }
            }
            catch (Exception)
            {
               
            }
            
        }       

        private void btnDeleteGroup_Click(object sender, EventArgs e)
        {
            try
            {
                if (treeData.DataSource == null) return;
                int id = (int)treeData.Selection[0].GetValue(colIDTree);
                if (id == 0) return;

                if (FormAndFunctionBO.Instance.CheckExist("FormAndFunctionGroupID", id))
                {
                    MessageBox.Show("Nhóm này đang được sử dụng nên không thể xóa được.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if (isParent(id))
                {
                    MessageBox.Show("Bạn phải xóa hết các nhóm con của nhóm này!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if (MessageBox.Show("Bạn có chắc muốn xóa nhóm [" + treeData.Selection[0].GetValue(colNameTree).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                FormAndFunctionGroupBO.Instance.Delete(id);
                loadTree();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Permission Events
        private void btnNew_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
            if(id >0 && !isParent(id))
            {
                frmPermission frm = new frmPermission();
                frm.CatID = id;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    LoadInfoSearch();
                }
            }            
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;
            FormAndFunctionModel model = (FormAndFunctionModel)FormAndFunctionBO.Instance.FindByPK(id);
            int catId = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colCat));
            _RownIndex = grvData.FocusedRowHandle;
            frmPermission frm = new frmPermission();
            frm.CatID = catId;
            frm.Model = model;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (grvData.DataSource == null) return;
                int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
                if (id == 0) return;

                if (UserRightDistributionBO.Instance.CheckExist("FormAndFunctionID", id))
                {
                    MessageBox.Show("Quyền này đang được sử dụng nên không thể xóa được.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if (MessageBox.Show("Bạn có chắc muốn xóa quyền [" + grvData.GetFocusedRowCellValue(colName).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                FormAndFunctionBO.Instance.Delete(id);
                LoadInfoSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        private void treeData_FocusedNodeChanged(object sender, FocusedNodeChangedEventArgs e)
        {
            try
            {
                int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
                if (id > 0 && !isParent(id))
                {
                    LoadInfoSearch();
                }
            }
            catch (Exception)
            {
            }         
        }

        private void treeData_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                TreeListNode dragNode, targetNode;
                TreeList tl = sender as TreeList;
                Point p = tl.PointToClient(new Point(e.X, e.Y));
                dragNode = e.Data.GetData(typeof(TreeListNode)) as TreeListNode;
                targetNode = tl.CalcHitInfo(p).Node;

                int targetID = TextUtils.ToInt(targetNode.GetValue(colIDTree));
                int drapID = TextUtils.ToInt(dragNode.GetValue(colIDTree));
                FormAndFunctionGroupModel model = (FormAndFunctionGroupModel)FormAndFunctionGroupBO.Instance.FindByPK(drapID);
                model.ParentID = targetID;
                FormAndFunctionGroupBO.Instance.Update(model);

                tl.SetNodeIndex(dragNode, tl.GetNodeIndex(targetNode));
                e.Effect = DragDropEffects.Move;
            }
            catch (Exception)
            {
            }
        }
    }
}
