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

namespace BMS
{
    public partial class frmPermission : _Forms
    {
        #region Variables
        public FormAndFunctionModel Model = new FormAndFunctionModel();
        public int CatID = 0;

        #endregion

        public frmPermission()
        {
            InitializeComponent();
        }

        private void frmPermission_Load(object sender, EventArgs e)
        {
            loadCombo(CatID);
            leParentCat.EditValue = CatID;
            if (Model.ID != 0)
            {
                txtName.Text = Model.Name;
                txtCode.Text = Model.Code;
            }
        }

        #region Functions

        void loadCombo(int id = 0)
        {
            DataTable tbl = TextUtils.Select(@"SELECT ID,Name as Name FROM FormAndFunctionGroup WITH(NOLOCK) ORDER BY Name");
            if (tbl == null)
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
                MessageBox.Show("Xin hãy điền Tên cho quyền này.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            if (txtCode.Text.Trim() == "")
            {                
                MessageBox.Show("Xin hãy điền Mã cho quyền này.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (Model.ID > 0)
                {
                     dt = TextUtils.Select("select Code from FormAndFunction where Code = '" + txtCode.Text.Trim() + "' and ID <> " + Model.ID );                   
                }
                else
                {
                     dt = TextUtils.Select("select Code from FormAndFunction where Code = '" + txtCode.Text.Trim() + "'");                   
                }
                if (dt != null)
                {
                    if (dt.Rows.Count>0)
                    {
                        MessageBox.Show("Mã này đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }                    
                }
            }
            return true;
        }
        #endregion

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
                    Model = new FormAndFunctionModel();
                }

                Model.Name = txtName.Text.Trim();
                Model.Code = txtCode.Text.Trim();
                Model.FormAndFunctionGroupID = TextUtils.ToInt(leParentCat.EditValue);
                Model.IsHide = false;

                if (Model.ID == 0)
                {
                    Model.CreatedDate = TextUtils.GetSystemDate();
                    Model.CreatedBy = Global.AppUserName;
                    Model.UpdatedDate = Model.CreatedDate;
                    Model.UpdatedBy = Global.AppUserName;
                    Model.ID = (int)pt.Insert(Model);
                }
                else
                {
                    Model.UpdatedDate = TextUtils.GetSystemDate();
                    Model.UpdatedBy = Global.AppUserName;
                    pt.Update(Model);
                }

                pt.CommitTransaction();
                
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pt.CloseConnection();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
