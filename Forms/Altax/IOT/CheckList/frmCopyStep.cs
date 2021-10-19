using DevExpress.XtraPrinting.Native;
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
    public partial class frmCopyStep : Form
    {
        public int GroupID;
        public int WorkingStepID;
        public string GroupName;
        public string StepName;
        public string StepCode;
        public string SortOrder;
        public frmCopyStep()
        {
            InitializeComponent();
        }
        void LoadCBTo()
        {
            DataTable dt = TextUtils.Select($"SELECT * FROM dbo.ProductGroup where ID <> {GroupID}");
            Check.DataSource = dt;
            Check.DisplayMember = "ProductGroupName";
            Check.ValueMember = "ID";
        }
        void LoadAll()
        {
            txtGroupName.Text = GroupName;
            txtWorkingStepName.Text = StepName;
        }
        private void frmCopyStep_Load(object sender, EventArgs e)
        {
            LoadCBTo();
            LoadAll();
            chkSelectAll.Checked = true;
        }
        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool value = chkSelectAll.Checked;
            for (int i = 0; i < Check.ItemCount; i++)
            {
                Check.SetItemCheckState(i, value ? CheckState.Checked : CheckState.Unchecked);
            }
        }
        private void btnCopy_Click(object sender, EventArgs e)
        {
            int count = Check.CheckedItems.Count;//.CheckedItemsCount;
            for (int i = 0; i < count; i++)
            {
                int productGroupID = TextUtils.ToInt(Check.CheckedItems[i]);
                
                TextUtils.ExcuteSQL(string.Format($"EXEC spCopyWorkingStepToOtherProductGroup 	@WorkingStepDescription=N'{StepName}' ,@WorkingStepSortOrder={SortOrder} ,@WorkingStepCode=N'{StepCode}' ,@WorkingStepID={WorkingStepID} ,@ProductGroupID={productGroupID} "));
            }
            MessageBox.Show("Copy OK!");
            this.Close();
        }
    }
}
