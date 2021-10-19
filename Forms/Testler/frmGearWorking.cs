using BMS;
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
    public partial class frmGearWorking : _Forms
    {
        public int GearID = 0;
        public frmGearWorking()
        {
            InitializeComponent();
        }

        private void frmGearWorking_Load(object sender, EventArgs e)
        {
            LoadStep();
        }

        private void LoadStep()
        {
            string[] paraName = new string[1];
            object[] paraValue = new object[1];
            paraName[0] = "@TextFilter";
            paraValue[0] = "";
            ArrayList arr = GearBO.Instance.GetListObject("spGetGearDataTestler", paraName, paraValue);
            grdData.DataSource = arr;
        }

        private void grvDetail_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {

        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            LoadGearWorking();
        }

        private void LoadGearWorking()
        {
            string HYP = TextUtils.ToString(grvData.GetFocusedRowCellValue(colHYP));
            DataTable dt = new DataTable();
            dt = TextUtils.LoadDataFromSP("spGetGearWorking_ByHypCode", "A"
                , new string[] { "@HYP" }
                , new object[] { HYP });
            grdDetail.DataSource = dt;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            grvDetail.FocusedRowHandle = -1;
            for (int i = 0; i < grvDetail.RowCount; i++)
            {
                int id = TextUtils.ToInt(grvDetail.GetRowCellValue(i, colID2));
                if (id == 0) continue;
                string value = TextUtils.ToString(grvDetail.GetRowCellValue(i, colValue));
                decimal min = TextUtils.ToDecimal(grvDetail.GetRowCellValue(i, colMin));
                decimal max = TextUtils.ToDecimal(grvDetail.GetRowCellValue(i, colMax));
                decimal defaultValue = TextUtils.ToDecimal(grvDetail.GetRowCellValue(i, colDefaultValue));
                string tansuat = TextUtils.ToString(grvDetail.GetRowCellValue(i, colTanSuat));
                string updateSql = string.Format("UPDATE dbo.GearWorking SET PeriodValue = '{0}' , MinValue = {1}, MaxValue = {2}, DefaultValue = {3}, TanSuat = '{4}'" +
                    " WHERE ID = {5}", value, min, max, defaultValue, tansuat, id); //giá trị decimal sau khi Format dấu . thành dấu , nên lệnh sql bị sai
                TextUtils.ExcuteSQL(updateSql);
            }
            MessageBox.Show("OK");
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmGearWorkingDetail frm = new frmGearWorkingDetail();
            frm.GearID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadGearWorking();
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            try
            {
                int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colID2));
                if (id == 0) return;
                GearWorkingModel model = (GearWorkingModel)GearWorkingBO.Instance.FindByPK(id);
                frmGearWorkingDetail frm = new frmGearWorkingDetail();
                int gearID = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colGearID));
                frm.GearWorking = model;
                frm.GearID = gearID;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    LoadGearWorking();
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colID2));
                if (id == 0) return;

                if (MessageBox.Show("Bạn có chắc muốn xóa mục cần kiểm tra [" + grvDetail.GetFocusedRowCellValue(colWorkingName).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                GearWorkingBO.Instance.Delete(id);
                grvDetail.DeleteSelectedRows();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
