using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Utils;
using BMS.Business;
using BMS.Model;

namespace BMS
{
    public partial class frmStaffManager : _Forms
    {
        private int _rownIndex = 0;

        public frmStaffManager()
        {
            InitializeComponent();            
        }

        private void frmStaffManager_Load(object sender, EventArgs e)
        {
            loadGrid();
        }

        void loadGrid()
        {
            //DataTable dt = TextUtils.Select("Select *,Case when Sex = 1 then N'Nam' else N'Nữ' end GioiTinh from vUserInfo WITH(NOLOCK)");
            DataTable dt = TextUtils.Select("Select *,Case when Sex = 1 then N'Nam' else N'Nữ' end GioiTinh from Users WITH(NOLOCK)");
            grdData.DataSource = null;
            grdData.DataSource = dt;
            if (_rownIndex >= grvData.RowCount)
                _rownIndex = 0;
            if (_rownIndex > 0)
                grvData.FocusedRowHandle = _rownIndex;
            grvData.SelectRow(_rownIndex);
            grvData.BestFitColumns();

            NhomPhong.DataSource = TextUtils.Select("SELECT * FROM UserGroup");
            NhomPhong.ValueMember = "ID";
            NhomPhong.DisplayMember = "Name";
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                TextUtils.ExportExcel(grvData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }     
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmShowStaff frm = new frmShowStaff();
            if (frm.ShowDialog()==DialogResult.OK)
            {
                loadGrid();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;

            _rownIndex = grvData.FocusedRowHandle;

            UsersModel model = (UsersModel)UsersBO.Instance.FindByPK(id);
            
            frmShowStaff frm = new frmShowStaff();
            frm.Model = model;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                loadGrid();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;
            string name = grvData.GetFocusedRowCellValue(colFullName).ToString();
            DialogResult result = MessageBox.Show("Bạn có thực sự muốn chuyển trạng thái của nhân viên [" + name + "]?", TextUtils.Caption,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                UsersModel model = (UsersModel)UsersBO.Instance.FindByPK(id);
                model.Status = 1;//ngừng hoạt động
                UsersBO.Instance.Update(model);
                loadGrid();
            }       
        }

        private void grdData_DoubleClick(object sender, EventArgs e)
        {
            btnEdit_Click(null, null);
        }

        

        private void btnQLNhom_Click(object sender, EventArgs e)
        {
            frmStaffGroup frm = new frmStaffGroup();
            frm.Show();
        }

        private void grvData_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.RowHandle>=0)
            {
                int status = TextUtils.ToInt(grvData.GetRowCellValue(e.RowHandle, colStatus));
                if (status == 1)
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
            }
        }

    }
}
