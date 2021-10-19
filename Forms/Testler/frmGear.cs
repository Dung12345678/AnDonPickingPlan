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
    public partial class frmGear : _Forms
    {
        public frmGear()
        {
            InitializeComponent();
        }

        private void frmProduct_Load(object sender, EventArgs e)
        {
            LoadInfoSearch();
            colGear3.VisibleIndex = 3;
            colLap.VisibleIndex = 4;
            colWipGCQT.VisibleIndex = 11;
        }

        private void LoadInfoSearch()
        {
            string[] paraName = new string[1];
            object[] paraValue = new object[1];

            paraName[0] = "@TextFilter"; paraValue[0] = txtName.Text.Trim();

            ArrayList arr = GearBO.Instance.GetListObject("spGetGearDataTestler", paraName, paraValue);
            grdData.DataSource = arr;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmGearDetail frm = new frmGearDetail();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadInfoSearch();
            }
        }
        private void btnCopy_Click(object sender, EventArgs e)
        {

        }

        int _rownIndex = 0;
        private void btnEdit_Click(object sender, EventArgs e)
        {
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;
            GearModel model = (GearModel)GearBO.Instance.FindByPK(id);
            int catId = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colWipGCQT));
            _rownIndex = grvData.FocusedRowHandle;

            frmGearDetail frm = new frmGearDetail();
            frm.GearModel = model;
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

            string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue(colHYP));

            //if (ProductCheckHistoryDetailBO.Instance.CheckExist("ProductID", strID))
            //{
            //    MessageBox.Show("Sản phẩm này đã có lịch sử kiểm tra nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //    return;
            //}

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa sản phẩm [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    GearBO.Instance.Delete(strID);
                    grvData.DeleteRow(grvData.FocusedRowHandle);
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
                }
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (grvData.RowCount > 0)
                { FolderBrowserDialog od = new FolderBrowserDialog();
                    if (od.ShowDialog() == DialogResult.OK)
                    {
                        TextUtils.ExportExcel(grvData,od.SelectedPath,string.Format("DanhSachSanPham_{0}",DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
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
