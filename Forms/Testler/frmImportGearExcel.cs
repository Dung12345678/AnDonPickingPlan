using BMS.Business;
using BMS.Model;
using BMS.Utils;
using DevExpress.XtraEditors;
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
	public partial class frmImportGearExcel : _Forms
	{
		public frmImportGearExcel()
		{
			InitializeComponent();
		}
		private void btnBrowse_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog();

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				btnBrowse.Text = ofd.FileName;
				cboSheet.DataSource = null;
				cboSheet.DataSource = TextUtils.ListSheetInExcel(ofd.FileName);
			}
		}

		private void frmImportGearExcel_Load(object sender, EventArgs e)
		{
            //ArrayList
		}

		private void cboSheet_SelectionChangeCommitted(object sender, EventArgs e)
		{
			try
			{
				DataTable dt = TextUtils.ExcelToDatatableNoHeader(btnBrowse.Text, cboSheet.SelectedValue.ToString());

				grdData.DataSource = dt;
				grvData.PopulateColumns();
				grvData.BestFitColumns();
                grdData.Focus();
               

            }
			catch (Exception ex)
			{
				TextUtils.ShowError(ex);
				grdData.DataSource = null;
			}
		}
		void enableControl(bool enable)
		{
			btnSave.Enabled = enable;
			grdData.Enabled = enable;
			cboSheet.Enabled = enable;
			btnBrowse.Enabled = enable;
		}

		DateTime start;
		private void btnSave_Click(object sender, EventArgs e)
		{
			if (backgroundWorker1.IsBusy)
			{
				backgroundWorker1.CancelAsync();
			}
			else
			{
				progressBar1.Minimum = 1;
				progressBar1.Maximum = grvData.RowCount;
				txtRate.Text = "";
				start = DateTime.Now;
				enableControl(false);
				backgroundWorker1.RunWorkerAsync();
			}
		}

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
                     
            int rowCount = grvData.RowCount;
            GearModel gear;
            GearWorkingModel gearWorking;
            int col = grvData.Columns.Count;
            int gearID;
            for (int i = 0; i < rowCount; i++)
            {
                if (i < 2) continue;
                progressBar1.Invoke((Action)(() => { progressBar1.Value = i - 1; }));
                txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i - 1, rowCount - 2); }));
                string HYPCode = Lib.ToString(grvData.GetRowCellValue(i, "F2"));
                Expression exp = new Expression("HYP", HYPCode);
                ArrayList arr = GearBO.Instance.FindByExpression(exp);
                if (arr.Count > 0)
                {
                    gear = arr[0] as GearModel;

                }
                else
                {
                    gear = new GearModel();
                }

                // update gear
                gear.HYP = Lib.ToString(grvData.GetRowCellValue(i, "F2"));
                gear.Gear1 = Lib.ToString(grvData.GetRowCellValue(i, "F3"));
                gear.Gear2 = Lib.ToString(grvData.GetRowCellValue(i, "F4"));
                gear.Gear3 = Lib.ToString(grvData.GetRowCellValue(i, "F5"));
                gear.Lap = Lib.ToString(grvData.GetRowCellValue(i, "F6"));
                gear.LotSize = Lib.ToInt(grvData.GetRowCellValue(i, "F7"));
                gear.GThieu = Lib.ToInt(grvData.GetRowCellValue(i, "F8"));
                gear.GGiaCongThem = Lib.ToInt(grvData.GetRowCellValue(i, "F9"));
                gear.WipG = Lib.ToInt(grvData.GetRowCellValue(i, "F10"));
                gear.QtyImage = Lib.ToInt(grvData.GetRowCellValue(i, "F11"));
                gear.QtyAudio = 2;

                if (arr.Count > 0)
                {
                    // update Gear                    
                    GearBO.Instance.Update(gear);
                    gearID = gear.ID;

                }
                else
                {
                    gearID=(int)GearBO.Instance.Insert(gear);

                }


                for (int j = 0; j < col; j++)
                {
                    string sortOrder = Lib.ToString(grvData.GetRowCellValue(1, grvData.Columns[j]));
                    if (string.IsNullOrWhiteSpace(sortOrder))
                        continue;
                    Expression exp1 = new Expression("gearID", gearID);
                    Expression exp2 = new Expression("SortOrder", TextUtils.ToInt(sortOrder));
                    ArrayList arrGW = GearWorkingBO.Instance.FindByExpression(exp1.And(exp2));
                    if (arrGW.Count > 0)
                    {
                        gearWorking = arrGW[0] as GearWorkingModel;
                        
                    }
                    else
                    {
                        gearWorking = new GearWorkingModel();
                    }
                    string value = Lib.ToString(grvData.GetRowCellValue(i, grvData.Columns[j]));
                    gearWorking.GearID = gearID;
                    gearWorking.MaxValue = TextUtils.ToDecimal(value);
                    gearWorking.MinValue = 0;
                    gearWorking.WorkingName = Lib.ToString(grvData.GetRowCellValue(0, grvData.Columns[j]));
                    gearWorking.CheckValueType = 1;
                    gearWorking.ValueType = 1;
                    gearWorking.SortOrder = TextUtils.ToInt(sortOrder);
                    gearWorking.TanSuat = "1/5";
                    gearWorking.DefaultValue = (gearWorking.MaxValue + gearWorking.MinValue) / 2;
                    if (value.Contains('-'))
                    {
                        decimal minValue = TextUtils.ToDecimal(value.Split('-')[0]);
                        decimal maxValue = TextUtils.ToDecimal(value.Split('-')[1]);
                        gearWorking.MaxValue = maxValue;
                        gearWorking.MinValue = minValue;
                        gearWorking.DefaultValue = (gearWorking.MaxValue + gearWorking.MinValue) / 2;

                    }

                    if (arrGW.Count > 0)
                    {
                        GearWorkingBO.Instance.Update(gearWorking);
                    }
                    else
                    {
                        GearWorkingBO.Instance.Insert(gearWorking);
                    }
                }                 
            }
        }

		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show(start.ToString() + " - " + DateTime.Now.ToString());
			enableControl(true);
		}
	}
}
