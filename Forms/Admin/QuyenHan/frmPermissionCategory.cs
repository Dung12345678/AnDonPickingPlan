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

namespace BMS
{
    public partial class frmPermissionCategory : _Forms
    {
        #region Variables
        public FormAndFunctionGroupModel Model =new FormAndFunctionGroupModel();
        public int ParentID = 0;
        public int CurentNode = 0;

        #endregion

        #region Contructors and Load
        public frmPermissionCategory()
        {
            InitializeComponent();
        }

        private void PriceDataCategory_Load(object sender, EventArgs e)
        {
            loadCombo(ParentID);
            leParentCat.EditValue = ParentID;
            if (Model.ID != 0)
            {
                txtName.Text = Model.Name;
            }
        }       
        #endregion

        #region Functions

        void loadCombo(int id = 0)
        {
            DataTable tbl = TextUtils.Select(@"SELECT ID,Name as Name FROM FormAndFunctionGroup WITH(NOLOCK) ORDER BY Name");
            if (tbl ==null)
            {
                return;
            }
            TextUtils.PopulateCombo(leParentCat, tbl.Copy(), "Name", "ID", "-- CẤP LỚN NHẤT --");
        }

        private void SetInterface(bool isNew)
        {            
            txtName.Enabled = isNew;
            leParentCat.Enabled = isNew;           
        }

        private bool ValidateForm()
        {          
            if (txtName.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền Tên cho nhóm quyền này.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            return true;
        }   
        #endregion

        #region Buttons Events
        private void btnSave_Click(object sender, EventArgs e)
        {
            ProcessTransaction pt = new ProcessTransaction();
            pt.OpenConnection();
            pt.BeginTransaction();
            try
            {
                if (!ValidateForm())
                    return;                

                if (Model == null)
                {
                    Model = new FormAndFunctionGroupModel();
                }
                Model.Name = txtName.Text.Trim();
                Model.ParentID = TextUtils.ToInt(leParentCat.EditValue);
                Model.IsHide = false;
                if (Model.ID == 0)
                {
                    Model.CreatedDate = TextUtils.GetSystemDate();
                    Model.CreatedBy = Global.AppUserName;
                    Model.UpdateDate = Model.CreatedDate;
                    Model.UpdatedBy = Global.AppUserName;
                    Model.ID = (int)pt.Insert(Model);
                }
                else
                {
                    Model.UpdateDate = TextUtils.GetSystemDate();
                    Model.UpdatedBy = Global.AppUserName;
                    pt.Update(Model);
                }

                pt.CommitTransaction();
                CurentNode = Model.ID;
                this.DialogResult = DialogResult.OK;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            { pt.CloseConnection(); }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region Events

        #endregion   

        private void treeData_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                //TreeListNode dragNode, targetNode;
                //TreeList tl = sender as TreeList;
                //Point p = tl.PointToClient(new Point(e.X, e.Y));
                //dragNode = e.Data.GetData(typeof(TreeListNode)) as TreeListNode;
                //targetNode = tl.CalcHitInfo(p).Node;

                //int targetID = TextUtils.ToInt(targetNode.GetValue(colIDTree));
                //int drapID = TextUtils.ToInt(dragNode.GetValue(colIDTree));
                //PriceDataCategoryModel model = (PriceDataCategoryModel)PriceDataCategoryBO.Instance.FindByPK(drapID);
                //model.ParentID = targetID;
                //PriceDataCategoryBO.Instance.Update(model);

                //tl.SetNodeIndex(dragNode, tl.GetNodeIndex(targetNode));
                //e.Effect = DragDropEffects.Move;
            }
            catch (Exception)
            {
               
            }
        }
    }
}
