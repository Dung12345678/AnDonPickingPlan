using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Model;
using BMS.Utils;
using BMS.Business;

namespace BMS
{
    public partial class frmChangePassword : _Forms
    {
        public frmChangePassword()
        {
            InitializeComponent();
        }

        private void frmChangePassword_Load(object sender, EventArgs e)
        {
            txtOldPass.Text = ((UsersModel)UsersBO.Instance.FindByPK(Global.UserID)).PasswordHash;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            UsersModel model = (UsersModel)UsersBO.Instance.FindByPK(Global.UserID);
            model.PasswordHash = MD5.EncryptPassword(txtNewPass.Text.Trim());
            UsersBO.Instance.Update(model);
            MessageBox.Show("Đổi mật khẩu thành công!","Thông báo");
            this.DialogResult = DialogResult.OK;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
