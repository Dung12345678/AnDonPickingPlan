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
    public partial class frmGearInfo : _Forms
    {
        public frmGearInfo()
        {
            InitializeComponent();
        }

        private void frmProduct_Load(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }

        private void LoadInfoSearch()
        {
            string sql = @"select *, -1 as Status from GearInfo";
            DataTable dt = TextUtils.Select(sql);
            grdData.DataSource = dt;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue(colGearCode));

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa mã hàng [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    GearInfoBO.Instance.Delete(strID);
                    grvData.DeleteRow(grvData.FocusedRowHandle);
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            grvData.FocusedRowHandle = -1;

            for (int i = 0; i < grvData.RowCount; i++)
            {
                int status = TextUtils.ToInt(grvData.GetRowCellValue(i, colStatus));
                if (status == -1) continue;

                int id = TextUtils.ToInt(grvData.GetRowCellValue(i, colID));
                string gearCode = TextUtils.ToString(grvData.GetRowCellValue(i, colGearCode));
                decimal sMin= TextUtils.ToDecimal(grvData.GetRowCellValue(i, colSlitMin));
                decimal sMax= TextUtils.ToDecimal(grvData.GetRowCellValue(i, colSlitMax));
                decimal vMin= TextUtils.ToDecimal(grvData.GetRowCellValue(i, colVibrateMin));
                decimal vMax= TextUtils.ToDecimal(grvData.GetRowCellValue(i, colVibrateMax));

                GearInfoModel model = new GearInfoModel();
                model.GearCode = gearCode;
                model.SlitMin = sMin;
                model.SlitMax = sMax;
                model.VibrateMax = vMax;
                model.VibrateMin = vMin;
                model.ID = id;

                if (id == 0)
                {
                    GearInfoBO.Instance.Insert(model);
                }
                else
                {
                    if(status == 2)
                    {
                        GearInfoBO.Instance.Update(model);
                    }
                }
            }

            LoadInfoSearch();
        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            LoadInfoSearch();
        }

        private void grvData_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.RowHandle < 0) return;
            if (e.Column == colStatus) return;

            grvData.SetRowCellValue(e.RowHandle, colStatus, 2);
        }
    }
}
