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
    public partial class frmChooseEmployee : _Forms
    {
        public bool IsChooseMulti = true;
        public List<int> LstID = new List<int>();

        public frmChooseEmployee()
        {
            InitializeComponent();
        }

        private void frmChooseEmployee_Load(object sender, EventArgs e)
        {
            grvDetail.OptionsSelection.MultiSelect = IsChooseMulti;
            loadDepartment();
            loadEmployee();
        }

        void loadDepartment()
        {
            DataTable dt = TextUtils.Select("SELECT * FROM dbo.Department WITH(NOLOCK)");
            cboDepartment.Properties.DataSource = dt;
            cboDepartment.Properties.DisplayMember = "Name";
            cboDepartment.Properties.ValueMember = "ID";
        }

        void loadEmployee()
        {
            int dID = TextUtils.ToInt(cboDepartment.EditValue);
            DataTable dt = TextUtils.GetDataTableFromSP("spGetEmployeeByDepartmentID", new string[] { "@DepartmentID" }, new object[] { dID });
            grdDetail.DataSource = dt;
        }

        private void cboDepartment_EditValueChanged(object sender, EventArgs e)
        {
            loadEmployee();
        }
        
        private void btnSave_Click(object sender, EventArgs e)
        {
            LstID = new List<int>();
            int[] lst = grvDetail.GetSelectedRows();
            for (int i = 0; i < lst.Length; i++)
            {
                LstID.Add(TextUtils.ToInt(grvDetail.GetRowCellValue(lst[i], colID)));
            }
            DialogResult = DialogResult.OK;
        }
    }
}
