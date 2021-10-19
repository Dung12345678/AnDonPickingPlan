using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMS.Business;
using IE.Model;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using GemBox.Spreadsheet;

namespace BMS
{
    public partial class frmExportAssemble : _Forms
    {
        public DateTime DateLR;
        string accessory = "";
        DataTable dtnew;
        DataTable dtLapRap;
        DataTable dtAccessory;
        int productID = 0;
        int countStepCode; // tổng cổng đoạn

        public frmExportAssemble()
        {
            InitializeComponent();
        }
        private void frmExportAssemble_Load(object sender, EventArgs e)
        {
            pnlLapRap.Visible = pnlLinhkienlaprap.Visible = btnExcelAssemble.Visible = false;
        }

        #region Event
        /// <summary>
        /// lấy ra danh sách lắp ráp
        /// </summary>
        void loadLapRap()
        {
            string stepCode = LibIE.ToString(cbStepCode.EditValue);
            //dtLapRap = LibIE.LoadDataFromSP("spGetHistoryCheckDataDetail", "A", new string[] { "@QRCode", "@ProductStepID" }, new object[] { txtQRCode.Text.Trim(), stepCode });
            dtLapRap = LibIE.Select($"SELECT CAST(ROW_NUMBER()OVER(PARTITION BY s.SortOrder ORDER BY s.SortOrder,PD.CreatedDate) as nvarchar(150)) AS RowNum,PD.*, CAST(s.SortOrder as varchar(20)) + ' - ' + PD.ProductStepCode + ' - ' + PD.ProductStepName + N'  -  Người lắp: ' + pD.WorkerCode + ' - ' + PD.Line AS ProductStepInfo, PD.ValueTypeName AS ValueTypeText FROM dbo.ProductCheckHistoryDetail PD left join ProductStep s on s.ID = pd.ProductStepID WHERE PD.QRCode = '{txtQRCode.Text}' AND(PD.ProductStepID IN({stepCode}))");
            grdLapRap.DataSource = dtLapRap;
        }

        /// <summary>
        /// lấy ra danh sách linh kiện lắp ráp
        /// </summary>
        void loadLinhKienLapRap()
        {
            txtAccessory.Text = accessory;
            DataTable dt = TextUtilsHP.LoadDataFromSP("spGetExportComponent", "A", new string[] { "@Code" }, new object[] { accessory });
            DataTable dtnew = dt.Clone();
            dtnew.Clear();
            DataTable dtCurrent = (DataTable)grdLinhKienLapRap.DataSource;
            if (dtCurrent != null) dtnew.Merge(dtCurrent);
            if (dt != null) dtnew.Merge(dt);
            grdLinhKienLapRap.DataSource = dtnew;
        }

        /// <summary>
        /// lấy ra danh sách công đoạn
        /// </summary>
        void loadStepCode()
        {
            if (productID == 0)
            {
                cbStepCode.Properties.DataSource = "";
                return;
            }
            DataTable dtStepHYP = LibIE.Select($"SELECT * FROM [dbo].[ProductStep] where ProductID = '{productID}'");
            DataTable dtStepTest = TextUtils.Select($"SELECT * FROM [dbo].[ProductStep] where ProductID = '{productID}'");
            if (dtStepHYP.Rows.Count > 0)
            {
                cbStepCode.Properties.DisplayMember = "ProductStepCode";
                cbStepCode.Properties.ValueMember = "ID";
                cbStepCode.Properties.DataSource = dtStepHYP;
                countStepCode = dtStepHYP.Rows.Count;
            }
            else if (dtStepTest.Rows.Count > 0)
            {
                cbStepCode.Properties.DisplayMember = "ProductStepCode";
                cbStepCode.Properties.ValueMember = "ID";
                cbStepCode.Properties.DataSource = dtStepTest;
                countStepCode = dtStepTest.Rows.Count;
            }
        }

        /// <summary>
        /// danh sách linh kiện
        /// </summary>
        void loadAccessory()
        {
            dtAccessory = TextUtilsStock.LoadDataFromSP("spGetAccessory", "A", new string[] { "@OrderCode" }, new object[] { txtOrderCode.Text.Trim() });
            cbAccessory.Properties.DisplayMember = "SumColumn";
            cbAccessory.Properties.ValueMember = "ID";
            cbAccessory.Properties.DataSource = dtAccessory;
        }
        #endregion

        private bool ValidateForm()
        {
            if (txtQRCode.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng nhập mã QR-Code!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (txtProductCode.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng kiểm tra mã hàng đang để trống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (txtOrderCode.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng kiểm tra số order đang để trống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (cbAccessory.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng kiểm tra danh sách linh kiện đang để trống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private void txtOrder_TextChanged(object sender, EventArgs e)
        {
            if (txtOrderCode.Text != "") loadAccessory();
        }

        #region Button Event
        /// <summary>
        /// Click button reset
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            txtQRCode.Clear();
            txtProductCode.Clear();
            txtOrderCode.Clear();
            pnlLapRap.Visible = pnlLinhkienlaprap.Visible = btnExcelAssemble.Visible = chkLapRap.Checked = chkLinhKienLapRap.Checked = false;
            cbAccessory.EditValue = -1;
            cbStepCode.Properties.DataSource = null;
            for (int i = grvLinhKienLapRap.RowCount - 1; i >= 0; i--)
            {
                grvLinhKienLapRap.DeleteRow(i);
            }
        }

        /// <summary>
        /// Click button xuất file Excel linh kiện lắp ráp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportExcelLinhKien_Click(object sender, EventArgs e)
        {
            if (!ValidateForm()) return;

            string path = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            else
            {
                return;
            }
            string fileSourceName = "ExportComponent.xlsx";

            string sourcePath = Application.StartupPath + "\\" + fileSourceName;
            string currentPath = path + "\\" + txtQRCode.Text.Trim() + ".xlsx";
            try
            {
                File.Copy(sourcePath, currentPath, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi tạo phiếu!" + Environment.NewLine + ex.Message,
                    TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang tạo phiếu..."))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Excel.Application app = default(Excel.Application);
                Excel.Workbook workBoook = default(Excel.Workbook);
                Excel.Worksheet workSheet = default(Excel.Worksheet);
                try
                {
                    app = new Excel.Application();
                    app.Workbooks.Open(currentPath);
                    workBoook = app.Workbooks[1];
                    workSheet = (Excel.Worksheet)workBoook.Worksheets[1];

                    workSheet.Cells[3, 2] = txtQRCode.Text.Trim();
                    workSheet.Cells[3, 13] = chkLinhKienLapRap.Text.Trim();
                    workSheet.Cells[5, 2] = txtProductCode.Text.Trim();
                    workSheet.Cells[6, 2] = txtOrderCode.Text.Trim();
                    workSheet.Cells[8, 1] = "Linh Kiện";

                    for (int i = grvLinhKienLapRap.RowCount - 1; i >= 0; i--)
                    {
                        workSheet.Cells[10, 1] = i + 1;
                        workSheet.Cells[10, 2] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colCodeFull));
                        workSheet.Cells[10, 3] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colCodeShort));
                        workSheet.Cells[10, 4] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue));
                        workSheet.Cells[10, 5] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue1));
                        workSheet.Cells[10, 6] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue2));
                        workSheet.Cells[10, 7] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue3));
                        workSheet.Cells[10, 8] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue4));
                        workSheet.Cells[10, 9] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue5));
                        workSheet.Cells[10, 10] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue6));
                        workSheet.Cells[10, 10] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue7));
                        workSheet.Cells[10, 12] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue8));
                        workSheet.Cells[10, 13] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue9));
                        workSheet.Cells[10, 14] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue10));
                        workSheet.Cells[10, 14] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue11));
                        workSheet.Cells[10, 16] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue12));
                        workSheet.Cells[10, 17] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue13));
                        workSheet.Cells[10, 18] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue14));
                        workSheet.Cells[10, 19] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue15));
                        workSheet.Cells[10, 20] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue16));
                        workSheet.Cells[10, 21] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(i, colRealValue17));
                        ((Excel.Range)workSheet.Rows[10]).Insert();
                    }
                    ((Excel.Range)workSheet.Rows[9]).Delete();
                    ((Excel.Range)workSheet.Rows[9]).Delete();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (app != null)
                    {
                        app.ActiveWorkbook.Save();
                        app.Workbooks.Close();
                        app.Quit();
                    }
                }
                Process.Start(currentPath);
            }
        }

        /// <summary>
        /// click button xóa dòng
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            grvLinhKienLapRap.DeleteSelectedRows();
        }

        /// <summary>
        /// CLICK button xuất file excel lắp ráp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportExcelLapRap_Click(object sender, EventArgs e)
        {
            if (grvLapRap.RowCount <= 0) return;
            string filePath = TextUtils.ExportExcelReturnFilePath(grvLapRap, txtQRCode.Text.Trim());
            if (string.IsNullOrEmpty(filePath)) return;

            Excel.Application app = default(Excel.Application);
            Excel.Workbook workBoook = default(Excel.Workbook);
            Excel.Worksheet workSheet = default(Excel.Worksheet);
            try
            {
                app = new Excel.Application();
                app.Workbooks.Open(filePath);
                workBoook = app.Workbooks[1];
                workSheet = (Excel.Worksheet)workBoook.Worksheets[1];
                app.DisplayAlerts = false;
                for (int i = 0; i < 5; i++)
                {
                    ((Excel.Range)workSheet.Rows[1]).Insert();
                }
                string[] arr = txtQRCode.Text.Split(' ');
                workSheet.Cells[1, 1] = "Order No:";
                workSheet.Cells[1, 2] = txtQRCode.Text;
                workSheet.Cells[2, 1] = "PID No:";
                workSheet.Cells[2, 2] = arr[1];

                ProductModel p = (ProductModel)ProductBO.Instance.FindByAttribute("ProductCode", arr[1].Trim())[0];

                workSheet.Cells[3, 1] = "Mô Tả";
                workSheet.Cells[3, 2] = p.ProductName;
                workSheet.Cells[4, 1] = "Ngày lắp:";
                workSheet.Cells[4, 2] = "'" + DateLR.ToString("dd/MM/yyyy");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                app.ActiveWorkbook.Save();
                app.Workbooks.Close();
                app.Quit();

                Process.Start(filePath);
            }
        }


        /// <summary>
        /// click xuất excel tổng 2 bảng
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcelAssemble_Click(object sender, EventArgs e)
        {
            //if (grvLapRap.RowCount <= 0) return;
            //SaveFile();

            if (grvLapRap.RowCount <= 0) return;
            string filePath = TextUtils.ExportExcelReturnFilePath(grvLapRap, txtQRCode.Text);
            if (string.IsNullOrEmpty(filePath)) return;

            Excel.Application app = default(Excel.Application);
            Excel.Workbook workBoook = default(Excel.Workbook);
            Excel.Worksheet workSheet = default(Excel.Worksheet);
            try
            {
                app = new Excel.Application();
                app.Workbooks.Open(filePath);
                workBoook = app.Workbooks[1];
                workSheet = (Excel.Worksheet)workBoook.Worksheets[1];
                app.DisplayAlerts = false;
                for (int i = 0; i < 6; i++)
                {
                    ((Excel.Range)workSheet.Rows[1]).Insert();
                }
                string[] arr = txtQRCode.Text.Split(' ');
                workSheet.Cells[1, 1] = "QR-Code";
                workSheet.Cells[1, 2] = txtQRCode.Text;
                workSheet.Cells[2, 1] = "PID No:";
                workSheet.Cells[2, 2] = arr[1];
                workSheet.Cells[3, 1] = "KT dữ liệu";
                workSheet.Cells[3, 2] = chkLapRap.Text;
                workSheet.Cells[4, 1] = "DS linh kiện:";
                workSheet.Cells[4, 2] = cbAccessory.Text;
                workSheet.Cells[5, 1] = "Linh kiện";
                workSheet.Cells[5, 2] = txtAccessory.Text;

                int number = grvLapRap.RowCount + 10;
                for (int j = number; j <= number + grvLinhKienLapRap.RowCount; j++)
                {
                    int k = j - number - 1;
                    // hiển thị dòng đầu header "HLM"
                    if (j == number)
                    {
                        workSheet.Cells[j, 1] = "Linh Kiện";
                        workSheet.Cells[j, 2] = "HLM";
                        workSheet.Cells[j, 3] = "HLM";
                        workSheet.Cells[j, 1].Style.Font.Size = workSheet.Cells[j, 2].Style.Font.Size = workSheet.Cells[j, 3].Style.Font.Size = 16;
                        workSheet.Cells[j, 2].Interior.Color = workSheet.Cells[j, 3].Interior.Color = Excel.XlRgbColor.rgbDarkOrange; // màu
                        workSheet.Cells[j, 1].Style.HorizontalAlignment = workSheet.Cells[j, 2].Style.HorizontalAlignment = workSheet.Cells[j, 3].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // căn giữa

                        for (int m = 4; m <= 21; m++)
                        {
                            workSheet.Cells[j, m] = "HLM";
                            workSheet.Cells[j, m].Style.Font.Size = 16;
                            workSheet.Cells[j, m].Interior.Color = Excel.XlRgbColor.rgbYellowGreen; // màu
                            workSheet.Cells[j, m].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // căn giữa
                        }
                    }
                    else
                    {
                        workSheet.Cells[j, 1] = k + 1;
                        workSheet.Cells[j, 2] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colCodeFull));
                        workSheet.Cells[j, 3] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colCodeShort));
                        workSheet.Cells[j, 3].NumberFormat = "000";
                        workSheet.Cells[j, 4] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue));
                        workSheet.Cells[j, 5] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue1));
                        workSheet.Cells[j, 6] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue2));
                        workSheet.Cells[j, 7] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue3));
                        workSheet.Cells[j, 8] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue4));
                        workSheet.Cells[j, 9] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue5));
                        workSheet.Cells[j, 10] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue6));
                        workSheet.Cells[j, 10] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue7));
                        workSheet.Cells[j, 12] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue8));
                        workSheet.Cells[j, 13] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue9));
                        workSheet.Cells[j, 14] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue10));
                        workSheet.Cells[j, 14] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue11));
                        workSheet.Cells[j, 16] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue12));
                        workSheet.Cells[j, 17] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue13));
                        workSheet.Cells[j, 18] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue14));
                        workSheet.Cells[j, 19] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue15));
                        workSheet.Cells[j, 20] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue16));
                        workSheet.Cells[j, 21] = TextUtilsHP.ToString(grvLinhKienLapRap.GetRowCellValue(k, colRealValue17));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                app.ActiveWorkbook.Save();
                app.Workbooks.Close();
                app.Quit();

                Process.Start(filePath);
            }
        }
        #endregion

        /// <summary>
        /// enter textbox QRCode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtQRCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                cbStepCode.Enabled = true;
                if (txtQRCode.Text != "")
                {
                    cbAccessory.EditValue = -1;
                    txtOrderCode.Text = "";
                    txtProductCode.Text = "";
                }
                DataTable dtHYP = LibIE.Select("SELECT * FROM [SumitomoHyp].[dbo].[ProductCheckHistoryDetail] WHERE QRCode = '" + txtQRCode.Text.Trim() + "'");
                DataTable dtTest = TextUtils.Select("SELECT * FROM [SumitomoTest].[dbo].[ProductCheckHistoryDetail] WHERE QRCode = '" + txtQRCode.Text.Trim() + "'");
                if (dtHYP.Rows.Count > 0)
                {
                    DataRow[] rows = dtHYP.Select();
                    txtProductCode.Text = rows[0]["ProductCode"].ToString();
                    txtOrderCode.Text = rows[0]["OrderCode"].ToString();
                    productID = LibIE.ToInt(rows[0]["ProductID"].ToString());
                    cbAccessory.Enabled = true;
                }
                else if (dtTest.Rows.Count > 0)
                {
                    DataRow[] rows = dtTest.Select();
                    txtProductCode.Text = rows[0]["ProductCode"].ToString();
                    txtOrderCode.Text = rows[0]["OrderCode"].ToString();
                    productID = TextUtils.ToInt(rows[0]["ProductID"].ToString());
                    cbAccessory.Enabled = true;
                }
                else productID = 0;
                loadAccessory();
                loadStepCode();
                loadLapRap();
            }
        }

        /// <summary>
        /// lựa chọn danh sách linh kiện
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbAccessory_EditValueChanged(object sender, EventArgs e)
        {
            if (cbAccessory.Text.Trim() == "") return;
            string[] name = cbAccessory.Text.Trim().Split(',');
            if (name.Length > 1)
                accessory = name[name.Length - 2] + name[name.Length - 1];
            else
                accessory = cbAccessory.Text.Trim();
            loadLapRap();
            loadLinhKienLapRap();
            ShowPanel();
        }

        /// <summary>
        /// check  dữ liệu lắp ráp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkLapRap_CheckedChanged(object sender, EventArgs e)
        {
            ShowPanel();
        }

        /// <summary>
        /// check linh kiện lắp ráp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkLinhKienLapRap_CheckedChanged(object sender, EventArgs e)
        {
            ShowPanel();
        }

        /// <summary>
        /// hàm hiển thị panel khi chọn checkbox
        /// </summary>
        void ShowPanel()
        {
            if (chkLapRap.Checked == true && chkLinhKienLapRap.Checked == true)
            {
                pnlLapRap.Visible = pnlLinhkienlaprap.Visible = btnExcelAssemble.Visible = true;
                splitContainer1.Panel1.Show();
                splitContainer1.Panel2.Show();
                btnExportExcelLinhKien.Visible = btnExportExcelLapRap.Visible = false;
            }
            else
            {
                pnlLapRap.Visible = pnlLinhkienlaprap.Visible = btnExcelAssemble.Visible = false;
                btnExportExcelLinhKien.Visible = btnExportExcelLapRap.Visible = true;
            }
            if (chkLapRap.Checked == true)
            {
                pnlLapRap.Visible = true;
                splitContainer1.Panel1Collapsed = false;
                splitContainer1.Panel1.Show();
            }
            else
            {
                splitContainer1.Panel1Collapsed = true;
                splitContainer1.Panel1.Hide();
            }
            if (chkLinhKienLapRap.Checked == true)
            {
                pnlLinhkienlaprap.Visible = true;
                splitContainer1.Panel2Collapsed = false;
                splitContainer1.Panel2.Show();
            }
            else
            {
                splitContainer1.Panel2Collapsed = true;
                splitContainer1.Panel2.Hide();
            }
        }

        //public bool SaveFile()
        //{
        //    try
        //    {
        //        SpreadsheetInfo.SetLicense("ELAP-G41W-CZA2-XNNC");

        //        string filePath = TextUtils.ExportExcelReturnFilePath(grvLapRap, txtQRCode.Text.Trim());
        //        string filePathTemplate = Path.Combine(Application.StartupPath, "Templates") + Path.DirectorySeparatorChar + "Mau.xls";
        //        File.Copy(filePathTemplate, filePath, true);

        //        if (string.IsNullOrEmpty(filePath)) return false;

        //        GemBox.Spreadsheet.ExcelFile spreadsheet = new GemBox.Spreadsheet.ExcelFile();
        //        ExcelFile workbook = ExcelFile.Load(filePath);
        //        ExcelWorksheet worksheet = workbook.Worksheets[0];

        //        int iStart = worksheet.NamedRanges["Start_Table1"].Range.FirstRowIndex;
        //        int iStart2 = worksheet.NamedRanges["Start_Table2"].Range.FirstRowIndex;

        //        worksheet.Cells[2, 2].Value = txtQRCode.Text.Trim();
        //        worksheet.Cells[2, 13].Value = chkLinhKienLapRap.Text.Trim();
        //        worksheet.Cells[4, 2].Value = txtProductCode.Text.Trim();
        //        worksheet.Cells[5, 2].Value = txtOrderCode.Text.Trim();
        //        worksheet.Cells[5, 12].Value = txtAccessory.Text.Trim();

        //        // bảng lắp ráp
        //        //worksheet.Cells[iStart, 1].Value = "TRUONG1";
        //        for (int i = dtLapRap.Rows.Count - 1; i >= 0; i--)
        //        {
        //            worksheet.Cells[iStart + 3, 0].Value = TextUtilsHP.ToString(dtLapRap.Rows[i]["RowNum"]);
        //            worksheet.Cells[iStart + 3, 1].Value = TextUtilsHP.ToString(dtLapRap.Rows[i]["ProductWorkingName"]);
        //            worksheet.Cells[iStart + 3, 7].Value = TextUtilsHP.ToString(dtLapRap.Rows[i]["ValueTypeText"]);
        //            worksheet.Cells[iStart + 3, 10].Value = TextUtilsHP.ToString(dtLapRap.Rows[i]["StandardValue"]);
        //            worksheet.Rows.InsertEmpty(iStart + 3);
        //        }


        //        // bảng danh sách linh kiện
        //        //worksheet.Cells[iStart2 + 2, 1].Value = "TRUONG2";
        //        for (int i = grvLinhKienLapRap.RowCount - 1; i >= 0; i--)
        //        {
        //            worksheet.Cells[iStart2 + 2, 0].Value = i + 1;
        //            worksheet.Cells[iStart2 + 2, 1].Value = TextUtilsHP.ToString(dtnew.Rows[i]["CodeFull"]);
        //            worksheet.Cells[iStart2 + 2, 2].Value = TextUtilsHP.ToString(dtnew.Rows[i]["CodeShort"]);
        //            worksheet.Cells[iStart2 + 2, 3].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue"]);
        //            worksheet.Cells[iStart2 + 2, 4].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue1"]);
        //            worksheet.Cells[iStart2 + 2, 5].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue2"]);
        //            worksheet.Cells[iStart2 + 2, 6].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue3"]);
        //            worksheet.Cells[iStart2 + 2, 7].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue4"]);
        //            worksheet.Cells[iStart2 + 2, 8].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue5"]);
        //            worksheet.Cells[iStart2 + 2, 9].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue6"]);
        //            worksheet.Cells[iStart2 + 2, 10].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue7"]);
        //            worksheet.Cells[iStart2 + 2, 11].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue8"]);
        //            worksheet.Cells[iStart2 + 2, 12].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue9"]);
        //            worksheet.Cells[iStart2 + 2, 13].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue10"]);
        //            worksheet.Cells[iStart2 + 2, 14].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue11"]);
        //            worksheet.Cells[iStart2 + 2, 15].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue12"]);
        //            worksheet.Cells[iStart2 + 2, 16].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue13"]);
        //            worksheet.Cells[iStart2 + 2, 17].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue14"]);
        //            worksheet.Cells[iStart2 + 2, 18].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue15"]);
        //            worksheet.Cells[iStart2 + 2, 19].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue16"]);
        //            worksheet.Cells[iStart2 + 2, 20].Value = TextUtilsHP.ToString(dtnew.Rows[i]["RealValue17"]);
        //            worksheet.Rows.InsertEmpty(iStart2 + 2);
        //        }

        //        workbook.Save(filePath);
        //        Process.Start(filePath);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}


        private void cbStepCode_EditValueChanged(object sender, EventArgs e)
        {
            loadLapRap();
            ShowPanel();
        }

        /// <summary>
        /// enter txtAccessory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtAccessory_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                DataTable dt = TextUtilsHP.LoadDataFromSP("spGetExportComponent", "A", new string[] { "@Code" }, new object[] { txtAccessory.Text });
                dtnew = dt.Clone();
                dtnew.Clear();
                DataTable dtCurrent = (DataTable)grdLinhKienLapRap.DataSource;
                if (dtCurrent != null) dtnew.Merge(dtCurrent);
                if (dt != null) dtnew.Merge(dt);
                grdLinhKienLapRap.DataSource = dtnew;
            }
        }
    }
}
