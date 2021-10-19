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

namespace BMS
{
    public partial class frmExportComponent : _Forms
    {
        public frmExportComponent()
        {
            InitializeComponent();
        }

        /// <summary>
        /// check lỗi
        /// </summary>
        /// <returns></returns>
        private bool ValidateForm()
        {
            if (txtQRCode.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng nhập mã QR-Code!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        /// <summary>
        /// enter txtQRCode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtQRCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                txtOrder.Text = "";
                txtArticleID.Text = "";
                if (txtQRCode.Text != "")
                {
                    DataTable dt = TextUtilsHP.LoadDataFromSP("spGetExportComponent", "A", new string[] { "@Code" }, new object[] { txtQRCode.Text.Trim() });
                    if (dt.Rows.Count == 0) return;
                    DataRow[] rows = dt.Select();
                    txtArticleID.Text = rows[0]["ArticleID"].ToString();
                    txtOrder.Text = rows[0]["OrderMachining"].ToString();
                    DataTable dtnew = dt.Clone();
                    dtnew.Clear();
                    DataTable dtCurrent = (DataTable)grdData.DataSource;
                    if(dtCurrent !=null) dtnew.Merge(dtCurrent);
                    if (dt != null) dtnew.Merge(dt);
                    grdData.DataSource = dtnew;
                }
            }
        }

        #region Buttons Event
        /// <summary>
        /// click button xuất excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportExcel_Click(object sender, EventArgs e)
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
                    workSheet.Cells[3, 11] = chkGiaCong.Text.Trim();
                    workSheet.Cells[5, 2] = txtArticleID.Text.Trim();
                    workSheet.Cells[6, 2] = txtOrder.Text.Trim();

                    for (int i = grvData.RowCount - 1; i >= 0; i--)
                    {
                        workSheet.Cells[10, 1] = "DỮ LIỆU KIỂM TRA";
                        workSheet.Cells[10, 2] = TextUtils.ToString(grvData.GetRowCellValue(i, colCodeFull));
                        workSheet.Cells[10, 3] = TextUtils.ToString(grvData.GetRowCellValue(i, colCodeShort));
                        workSheet.Cells[10, 4] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue));
                        workSheet.Cells[10, 5] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue1));
                        workSheet.Cells[10, 6] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue2));
                        workSheet.Cells[10, 7] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue3));
                        workSheet.Cells[10, 8] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue4));
                        workSheet.Cells[10, 9] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue5));
                        workSheet.Cells[10, 10] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue6));
                        workSheet.Cells[10, 11] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue7));
                        workSheet.Cells[10, 12] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue8));
                        workSheet.Cells[10, 13] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue9));
                        workSheet.Cells[10, 14] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue10));
                        workSheet.Cells[10, 15] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue11));
                        workSheet.Cells[10, 16] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue12));
                        workSheet.Cells[10, 17] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue13));
                        workSheet.Cells[10, 18] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue14));
                        workSheet.Cells[10, 19] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue15));
                        workSheet.Cells[10, 20] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue16));
                        workSheet.Cells[10, 21] = TextUtils.ToString(grvData.GetRowCellValue(i, colRealValue17));
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
            grvData.DeleteSelectedRows();
        }
		#endregion

		private void frmExportComponent_Load(object sender, EventArgs e)
		{

		}
	}
}
