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
    public partial class frmUserGroupDetail : _Forms
    {
        public UserGroupModel UserGroup = new UserGroupModel();
        public frmUserGroupDetail()
        {
            InitializeComponent();
        }

        private void frmUserGroupDetail_Load(object sender, EventArgs e)
        {
            loadData();
        }

        /// <summary>
        /// Load dữ liệu vào form
        /// </summary>
        void loadData()
        {
            txtCode.Text = UserGroup.Code;
            txtName.Text = UserGroup.Name;
        }

        /// <summary>
        /// Validate trước khi cất dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (txtCode.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền mã.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                if (TextUtils.CheckExistTable(UserGroup.ID, "Code", txtCode.Text.Trim(), "UserGroup"))
                {
                    MessageBox.Show("Mã này đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return false;
                }
            }
            if (txtName.Text.Trim() == "")
            {
                MessageBox.Show("Xin hãy điền tên.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            
            return true;
        }

        /// <summary>
        /// Cất dữ liệu
        /// </summary>
        /// <returns></returns>
        bool saveData()
        {
            if (!ValidateForm())
                return false;

            UserGroup.Code = txtCode.Text.Trim();
            UserGroup.Name = txtName.Text.Trim();

            if (UserGroup.ID > 0)
            {
                UserGroupBO.Instance.Update(UserGroup);
            }
            else
            {
                UserGroup.ID = (int)UserGroupBO.Instance.Insert(UserGroup);
            }
            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (saveData())
                this.DialogResult = DialogResult.OK;
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            if (saveData())
            {
                UserGroup = new UserGroupModel();
                loadData();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmUserGroupDetail_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
