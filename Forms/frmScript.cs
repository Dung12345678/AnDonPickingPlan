using BMS.Business;
using BMS.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmScript : _Forms
    {
        
        public frmScript()
        {
            InitializeComponent();
        }

        private void frmScript_Load(object sender, EventArgs e)
        {
           
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            TextUtils.ExcuteSQL(string.Format(@"{0}", txtName.Text.Trim()));
            txtName.Text = "";
        }
    }
}
