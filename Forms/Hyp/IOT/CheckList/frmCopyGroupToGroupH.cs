using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmCopyAll : Form
    {
        public int GroupID;
        public frmCopyAll()
        {
            InitializeComponent();
        }
        void LoadCBTo()
        {
            DataTable dt = LibIE.Select($"SELECT * FROM dbo.ProductGroup where ID <> {GroupID}");
            Check.DataSource = dt;
            Check.DisplayMember = "ProductGroupName";
            Check.ValueMember = "ID";
        }
        void LoadCBFrom()
        {
            DataTable dt = LibIE.Select($"SELECT * FROM dbo.ProductGroup where ID = {GroupID}");
            cboGroup.Properties.DataSource = dt;
            cboGroup.Properties.DisplayMember = "ProductGroupName";
            cboGroup.Properties.ValueMember = "ID";
            cboGroup.EditValue = TextUtils.ToString(dt.Rows[0]["ID"]);
        }
        private void frmCopyAll_Load(object sender, EventArgs e)
        {
            LoadCBTo();
            LoadCBFrom();
            Check.CheckAll();
           
        }
            
        private void btnCopy_Click(object sender, EventArgs e)
        {

            string GroupIDNew = TextUtils.ToString(Check.ValueMember);
            string[] arrListStr = GroupIDNew.Split(',');
            for (int i = 0; i < arrListStr.Length; i++)
            {
                LibIE.ExcuteSQL(string.Format($"EXEC spCopyALLGroupProducts @GroupID={GroupID}, @GroupNewID={arrListStr[i]}")); 
            }
        }
    }
}
