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
    public partial class frmPlan : _Forms
    {
        public frmPlan()
        {
            InitializeComponent();
        }

        private void frmProduct_Load(object sender, EventArgs e)
        {
            dtpFromDate.Value = DateTime.Now.AddDays(-30);
            dtpEndDate.Value = DateTime.Now;

            LoadInfoSearch();
        }

        private void LoadInfoSearch()
        {
            DateTime from = new DateTime(dtpFromDate.Value.Year, dtpFromDate.Value.Month, dtpFromDate.Value.Day, 0, 0, 0);
            DateTime to = new DateTime(dtpEndDate.Value.Year, dtpEndDate.Value.Month, dtpEndDate.Value.Day, 23, 59, 59);
         
            DataTable dt = TextUtils.LoadDataFromSP("spGetPlanDataTestler", "A"
                , new string[] { "@FromDate", "@EndDate", "@TextFilter" }
                , new object[] { from, to, txtName.Text.Trim() });
            grdData.DataSource = dt;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmPlanDetail frm = new frmPlanDetail();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }

        int _rownIndex = 0;
        private void btnEdit_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;
            PlanOrderModel model = (PlanOrderModel)PlanOrderBO.Instance.FindByPK(id);
            _rownIndex = grvData.FocusedRowHandle;

            frmPlanDetail frm = new frmPlanDetail();
            frm.PlanOrderModel = model;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));

            string strOrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
            string strGear = TextUtils.ToString(grvData.GetFocusedRowCellValue(colGearCode));

            //if (ProductCheckHistoryDetailBO.Instance.CheckExist("ProductID", strID))
            //{
            //    MessageBox.Show("Sản phẩm này đã có lịch sử kiểm tra nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return;
            //}

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa kế hoạch [{0}-{1}] không?", strOrderCode, strGear), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    PlanOrderBO.Instance.Delete(strID);
                    grvData.DeleteRow(grvData.FocusedRowHandle);
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
                }
                string[] arr = new string[5];
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (grvData.RowCount > 0)
                {
                    FolderBrowserDialog od = new FolderBrowserDialog();
                    if (od.ShowDialog() == DialogResult.OK)
                    {
                        TextUtils.ExportExcel(grvData,od.SelectedPath,string.Format("DanhSachKeHoach_{0}",DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void btnFindAll_Click(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }        

        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadInfoSearch();
            }
        }

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0 && btnEdit.Enabled == true)
                btnEdit_Click(null, null);
        } 

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            frmImportProductExcel frm = new frmImportProductExcel();
            frm.Show();
        }
    }
}
