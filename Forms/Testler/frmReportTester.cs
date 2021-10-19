using BMS;
using BMS.Business;
using BMS.Model;
using DevExpress.CodeParser;
using DevExpress.XtraGrid.Columns;
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
    public partial class frmReportTester : _Forms
    {
        public frmReportTester()
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
        }

        private void bgWorkerExportReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Enabled = true;
        }

        private void bgWorkerExportReport_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new MethodInvoker(exportExcel));
        }

        private void bgWorkerLoadReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Enabled = true;
        }

        private void bgWorkerLoadReport_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new MethodInvoker(getReport));
        }

        private void InitDataToComboBox()
        {
            
        }

        private void getReport()
        {
            //check valid period time
            if (DateTime.Parse(dtpFrom.Text) > DateTime.Parse(dtpTo.Text))
            {
                MessageBox.Show("Khoảng thời gian không hợp lệ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.Enabled = false;

            //get data report
            string from = new DateTime(dtpFrom.Value.Year, dtpFrom.Value.Month, dtpFrom.Value.Day, 0, 0, 0).ToString("yyyy-MM-dd HH:mm:ss");
            string to = new DateTime(dtpTo.Value.Year, dtpTo.Value.Month, dtpTo.Value.Day, 23, 59, 59).ToString("yyyy-MM-dd HH:mm:ss");
            string Hyp = txtHyp.Text;
            string Gear = txtGear.Text;
            string Lot = txtLot.Text;
            System.Data.DataTable dtReport = TextUtils.GetDataTableFromSP("spGetReportTester",
                new string[] { "@BeginDate", "@ToDate", "@Hyp", "@Gear", "@LotGiaCong" },
                new object[] { from, to, Hyp, Gear, Lot });

            if (dtReport != null)
            {
                grvData.Columns.Clear();
                GridColumn colQtyLap = new GridColumn();
                for (int i = 0; i < dtReport.Columns.Count; i++)
                {
                    GridColumn col = new GridColumn();
                    switch (dtReport.Columns[i].Caption)
                    {
                        case "DateReport":
                            col.Caption = "Ngày/ Tháng/ Năm";
                            col.Name = "colDateReport";
                            break;
                        case "WorkerCode":
                            col.Caption = "Người thao tác";
                            col.Name = "colWorkerCode";
                            break;
                        case "LotGiaCong":
                            col.Caption = "Lot gia công";
                            col.Name = "colLot";
                            break;
                        case "HypCode":
                            col.Caption = "Mã HYP";
                            col.Name = "colHyp";
                            break;
                        case "GearCode":
                            col.Caption = "Mã Gear";
                            col.Name = "colGear";
                            break;
                        case "QtyLap":
                            col.Caption = "Số lần Láp";
                            col.Name = "colQtyLap";
                            colQtyLap = col;
                            break;
                        default:
                            col.Caption = dtReport.Columns[i].Caption;
                            col.Name = "col" + (i + 1);
                            break;
                    }
                    if (col.Name.Equals("colQtyLap"))
                        continue;
                    col.VisibleIndex = i + 1;
                    col.FieldName = dtReport.Columns[i].Caption;
                    grvData.Columns.Add(col);
                }
                colQtyLap.VisibleIndex = grvData.Columns.Count + 1;
                colQtyLap.FieldName = "QtyLap";
                grvData.Columns.Add(colQtyLap);
            }

            grdReport.DataSource = dtReport;
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
                Workbook wb = xlApp.Workbooks.Open((sampleFile + @"\REPORT TESTER.xlsx").Trim());

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

                Range row3_Gear_value = ws.get_Range("B3", "C3");//Cột B, C dòng 3
                row3_Gear_value.Value2 = txtGear.Text;

                Range row3_Lot_value = ws.get_Range("E3");//Cột E dòng 3
                row3_Lot_value.Value2 = txtLot.Text;

                Range row4_Hyp_value = ws.get_Range("B4", "C4");//Cột B, C dòng 4
                row4_Hyp_value.Value2 = txtHyp.Text;

                //Header
                dynamic[] arrHeader = new dynamic[grvData.Columns.Count];
                for (int i = 0; i < grvData.Columns.Count; i++)
                {
                    arrHeader.SetValue(grvData.Columns[i].Caption, i);
                }
                Range rowHeader = ws.get_Range("A5", ColumnIndexToColumnLetter(grvData.Columns.Count) + "5");
                rowHeader.Value2 = arrHeader;
                rowHeader.EntireRow.Font.Bold = true;
                Range sampleGeneralHeaderCell = (Range)ws.Cells[5, 1];
                sampleGeneralHeaderCell.Copy(Type.Missing);
                rowHeader.Cells.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                //Dữ liệu báo cáo
                int dataRow = 8;
                Range sampleGeneralDataCell = (Range)ws.Cells[6, 2];
                for (int i = 0; i < grvData.RowCount; i++)
                {
                    dataRow++;
                    dynamic[] arrData = new dynamic[grvData.Columns.Count];
                    for (int j = 0; j < grvData.Columns.Count; j++)
                    {
                        string value = grvData.GetRowCellDisplayText(i, grvData.Columns[j]);
                        arrData.SetValue(value, j);
                    }
                    Range rowData = ws.get_Range("A" + dataRow, ColumnIndexToColumnLetter(grvData.Columns.Count) + dataRow);//Lấy dòng thứ row ra để đổ dữ liệu
                    rowData.Value2 = arrData;
                    sampleGeneralDataCell.Copy(Type.Missing);
                    rowData.Cells.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    Range cellDate = rowData.get_Range("A1");
                    cellDate.NumberFormat = "DD/MM/YYYY";
                }

                //Xóa 2 dòng mẫu
                for (int i = 0; i < 3; i++)
                {
                    Range SampleRow6 = ws.Rows[6];
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

        private string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }
}
