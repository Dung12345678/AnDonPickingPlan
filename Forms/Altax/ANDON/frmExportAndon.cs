using BMS;
using BMS.Business;
using BMS.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmExportAndon : _Forms
    {
        public frmExportAndon()
        {
            InitializeComponent();
        }

        BackgroundWorker bgWorkerLoadReport = new BackgroundWorker();
        BackgroundWorker bgWorkerExportReport = new BackgroundWorker();

        private void frmExportAndon_Load(object sender, EventArgs e)
        {
            InitDataToComboBox();
            bgWorkerLoadReport.DoWork += bgWorkerLoadReport_DoWork;
            bgWorkerLoadReport.RunWorkerCompleted += bgWorkerLoadReport_RunWorkerCompleted;
            bgWorkerExportReport.DoWork += bgWorkerExportReport_DoWork;
            bgWorkerExportReport.RunWorkerCompleted += bgWorkerExportReport_RunWorkerCompleted;
            cbbType.SelectedIndex = 2;
        }

        private void bgWorkerExportReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pgbExport.Visible = false;
            this.Enabled = true;
        }

        private void bgWorkerExportReport_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new MethodInvoker(exportExcel));
        }

        private void bgWorkerLoadReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pgbExport.Visible = false;
            this.Enabled = true;
        }

        private void bgWorkerLoadReport_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new MethodInvoker(getReport));
        }

        private void InitDataToComboBox()
        {
            //Init data to cbbStep
            ArrayList arrStep = new ArrayList();
            arrStep.Add("---Tất cả---");
            for (int i = 0; i < 9; i++)
            {
                arrStep.Add("CD" + (i + 1));
            }
            cbbProductStepCode.DataSource = arrStep;
            //Init data to cbbGroupProductCode
            ArrayList arrProductGroup = ProductGroupBO.Instance.FindAll();
            ProductGroupModel firstPosition = new ProductGroupModel();
            firstPosition.ProductGroupCode = "0";
            firstPosition.ProductGroupName = "---Tất cả---";
            arrProductGroup.Insert(0, firstPosition);
            cbbGroupProductCode.DataSource = arrProductGroup;
            cbbGroupProductCode.ValueMember = "ProductGroupCode";
            cbbGroupProductCode.DisplayMember = "ProductGroupName";
        }

        private void getReport()
        {
            if (cbbType.SelectedIndex < 0)
            {
                MessageBox.Show("Bạn chưa chọn loại báo cáo!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //check valid period time
            if (DateTime.Parse(dtpFrom.Text) > DateTime.Parse(dtpTo.Text))
            {
                MessageBox.Show("Khoảng thời gian không hợp lệ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            pgbExport.Visible = true;
            pgbExport.Style = ProgressBarStyle.Marquee;
            pgbExport.Left = this.Width / 2;
            pgbExport.Top = this.Height / 2;
            this.Enabled = false;

            //get data report
            string from = new DateTime(dtpFrom.Value.Year, dtpFrom.Value.Month, dtpFrom.Value.Day, 0, 0, 0).ToString("yyyy-MM-dd HH:mm:ss");
            string to = new DateTime(dtpTo.Value.Year, dtpTo.Value.Month, dtpTo.Value.Day, 23, 59, 59).ToString("yyyy-MM-dd HH:mm:ss");
            string productStepCode;
            string productGroupCode;

            if (cbbProductStepCode.SelectedIndex != 0)
                productStepCode = cbbProductStepCode.Text;
            else
                productStepCode = string.Empty;
            if (cbbGroupProductCode.SelectedIndex != 0)
                productGroupCode = cbbGroupProductCode.SelectedValue.ToString();
            else
                productGroupCode = string.Empty;
            int Type = cbbType.SelectedIndex;
            System.Data.DataTable dtReport = TextUtils.GetDataTableFromSP("spGetAndonReport",
                new string[] { "@StartDate", "@EndDate", "@ProductStepCode", "@ProductGroupCode", "@Type" },
                new object[] { from, to, productStepCode, productGroupCode, Type });
            int totalDelay = 0;
            int totalRisk = 0;
            int totalTimeDelay = 0;
            int totalTimeRisk = 0;

            totalRisk = TextUtils.ToInt(dtReport.Compute("Count(Type)", "Type = 2"));
            totalDelay = TextUtils.ToInt(dtReport.Compute("Count(Type)", "Type = 1"));

            totalTimeDelay = TextUtils.ToInt(dtReport.Compute("Sum(DelayTime)", "Type = 1"));
            totalTimeRisk = TextUtils.ToInt(dtReport.Compute("Sum(RiskTime)", "Type = 2"));

            grdReport.DataSource = dtReport;

            txtTotalDelay.Text = totalDelay.ToString();
            txtTotalRisk.Text = totalRisk.ToString();
            txtTimeDelay.Text = totalTimeDelay.ToString();
            txtTimeRisk.Text = totalTimeRisk.ToString();
        }

        private void btnGetReport_Click(object sender, EventArgs e)
        {
            bgWorkerLoadReport.RunWorkerAsync();
        }

        private void btnWriteExcel_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtReport = (System.Data.DataTable)grdReport.DataSource;
            if (dtReport == null || dtReport.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu", "Cảnh Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            bgWorkerExportReport.RunWorkerAsync();
        }

        private void exportExcel()
        {
            try
            {
                //Lấy thư mục lưu trữ
                string savefilePath = "";
                sfdSaveDialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";
                sfdSaveDialog.ShowDialog();
                savefilePath = sfdSaveDialog.FileName;
                if (string.IsNullOrEmpty(savefilePath))
                {
                    return;
                }

                pgbExport.Visible = true;
                pgbExport.Style = ProgressBarStyle.Marquee;
                pgbExport.Left = this.Width / 2;
                pgbExport.Top = this.Height / 2;
                this.Enabled = false;

                //Tạo app excel
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Lỗi không thể sử dụng được thư viện EXCEL");
                    return;
                }
                xlApp.Visible = false;
                object misValue = System.Reflection.Missing.Value;

                //Đọc file báo cáo mẫu
                string sampleFile = Directory.GetCurrentDirectory();
                Workbook wb = xlApp.Workbooks.Open(sampleFile + @"\Data ANDON Altax assy line.xlsx");

                Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);

                if (ws == null)
                {
                    MessageBox.Show("Không thể tạo được WorkSheet");
                    return;
                }

                //Dữ liệu filter
                Range row2_From_Date = ws.get_Range("B2");//Cột B dòng 2
                row2_From_Date.Value2 = dtpFrom.Value.ToString("dd/MM/yyyy");

                Range row2_To_Date = ws.get_Range("D2");//Cột D dòng 2
                row2_To_Date.Value2 = dtpTo.Value.ToString("dd/MM/yyyy");

                Range row3_CD_value = ws.get_Range("B3", "C3");//Cột B, C dòng 3
                row3_CD_value.Value2 = cbbProductStepCode.Text;

                Range row3_totalDelay_value = ws.get_Range("E3");//Cột E dòng 3
                row3_totalDelay_value.Value2 = txtTotalDelay.Text;

                Range row3_TimeDelay_value = ws.get_Range("G3");//Cột G dòng 3
                row3_TimeDelay_value.Value2 = txtTimeDelay.Text;

                Range row3_Group_value = ws.get_Range("B4", "C4");//Cột B, C dòng 4
                row3_Group_value.Value2 = cbbGroupProductCode.Text;

                Range row3_totalRisk_value = ws.get_Range("E4");//Cột E dòng 4
                row3_totalRisk_value.Value2 = txtTotalRisk.Text;

                Range row3_TimeRisk_value = ws.get_Range("G4");//Cột G dòng 4
                row3_TimeRisk_value.Value2 = txtTimeRisk.Text;

                //Dữ liệu báo cáo
                int dataRow = 7;
                for (int i = 0; i < grvData.RowCount; i++)
                {
                    dataRow++;
                    string datetimeValue = grvData.GetRowCellDisplayText(i, colDateTime);
                    string dateReport = datetimeValue.Substring(3, 2) + "/" + datetimeValue.Substring(0, 2) + "/" + datetimeValue.Substring(6, 4);
                    string ShiftName = grvData.GetRowCellDisplayText(i, colShift);
                    string ProductCode = grvData.GetRowCellDisplayText(i, colProductCode);
                    string ProductName = grvData.GetRowCellDisplayText(i, colProductName);
                    string OrderCode = grvData.GetRowCellDisplayText(i, colOrderCode);
                    string QrCode = grvData.GetRowCellDisplayText(i, colSerial);
                    string ProductStepCode = grvData.GetRowCellDisplayText(i, colStepCode);
                    string WorkerCode = grvData.GetRowCellDisplayText(i, colWorker);
                    string MakeTime = grvData.GetRowCellDisplayText(i, colMakeTime);
                    string DelayTime = grvData.GetRowCellDisplayText(i, colDelayTime);
                    string RiskTime = grvData.GetRowCellDisplayText(i, colRiskTime);
                    string RiskDescription = grvData.GetRowCellDisplayText(i, colRiskDescription);
                    dynamic[] arr = { dateReport, ShiftName, ProductCode, ProductName, OrderCode, QrCode, ProductStepCode, WorkerCode, MakeTime, DelayTime, RiskTime, RiskDescription };
                    Range rowData = ws.get_Range("A" + dataRow, "L" + dataRow);//Lấy dòng thứ row ra để đổ dữ liệu
                    rowData.Value2 = arr;
                    Range sampleDataCell = (Range) ws.Cells[6, 1];
                    sampleDataCell.Copy(Type.Missing);
                    rowData.Cells.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }

                //Xóa 2 dòng mẫu
                for (int i = 0; i < 2; i++)
                {
                    Range SampleRow6 = ws.get_Range("A6", "L6");
                    SampleRow6.Delete();
                }

                Range focusDataCell = (Range)ws.Cells[6, 1];
                focusDataCell.Select();

                //Lưu file excel xuống Ổ cứng
                wb.SaveAs(savefilePath);

                //đóng file để hoàn tất quá trình lưu trữ
                wb.Close(true, misValue, misValue);
                //thoát và thu hồi bộ nhớ cho COM
                xlApp.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(xlApp);

                //Mở File excel sau khi Xuất thành công
                System.Diagnostics.Process.Start(savefilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Hàm thu hồi bộ nhớ cho COM Excel
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                obj = null;
            }
            finally
            { GC.Collect(); }
        }
    }
}
