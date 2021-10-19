using DevExpress.XtraEditors.Controls;
using DevExpress.XtraExport;
using Org.BouncyCastle.Bcpg;
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
    public partial class frmCopyWorkingToGroup : Form
    {
        public int GroupID;
        public int WorkingStepID;
        public int WorkingID;
        public string GroupName;
        public string WorkingStepCode;
        public int STT;
        public string WorkingName;
        public frmCopyWorkingToGroup()
        {
            InitializeComponent();
        }

        private void frmCopy_Load(object sender, EventArgs e)
        {
            loadProductGroup();
            loadData();
            chkSelectAll.Checked = true;
        }
        void loadProductGroup()
        {
            DataTable dt = TextUtils.Select($"SELECT * FROM dbo.ProductGroup where ID <> {GroupID}");
            Check.DataSource = dt;
            Check.DisplayMember = "ProductGroupName";
            Check.ValueMember = "ID";
        }
        void loadData()
        {
            txtGroupName.Text = GroupName;
            txtStepName.Text = WorkingStepCode;
            txtSTT.Text = TextUtils.ToString(STT);
            txtWorkingName.Text = WorkingName;
        }
        private void btnCopy_Click(object sender, EventArgs e)
        {
            int count = Check.CheckedItems.Count;//.CheckedItemsCount;
            for (int i = 0; i < count; i++)
            {
                int productGroupID = TextUtils.ToInt(Check.CheckedItems[i]);
                TextUtils.ExcuteSQL(string.Format($"EXEC spCopyWorkingToOtherProductGroup " +
                    $"@WorkingSortOrder={STT} ," +
                    $"@WorkingID={WorkingID} ," +
                    $"@WorkingStepCode=N'{WorkingStepCode}' ," +
                    $"@ProductGroupID={productGroupID}"));
            }

            MessageBox.Show("Copy OK!");

            this.Close();
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool value = chkSelectAll.Checked;
            for (int i = 0; i < Check.ItemCount; i++)
            {
                Check.SetItemCheckState(i, value ? CheckState.Checked : CheckState.Unchecked);
            }


            //item.CheckState = value? CheckState.Checked:CheckState.Unchecked;

        }
    }
}
