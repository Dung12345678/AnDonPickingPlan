using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMS
{
	public partial class frmHistoryCheckOrderDetailH : _Forms
	{
		public string _QRCode;
		public DateTime _DateLR;
		public string _OrderCode;
		public frmHistoryCheckOrderDetailH()
		{
			InitializeComponent();
		}
		private void frmHistoryCheckOrderDetailH_Load(object sender, EventArgs e)
		{
			loadData();
		}
		void loadData()
		{
			DataTable dt = LibIE.LoadDataFromSP("spGetHistoryCheckOrderDataDetailH", "A"
				, new string[] { "orderCode" }
				, new object[] { _OrderCode });
			int count = dt.Columns.Count;
			if (count > 9)
			{
				string RealValue1 = Lib.ToString(dt.Columns[9]);
				colRealValue1.FieldName = RealValue1;
				colRealValue1.Caption = RealValue1;
			}
			else
				colRealValue1.Visible = false;
			if (count > 10)
			{
				string realvalue2 = Lib.ToString(dt.Columns[10]);
				colRealValue2.FieldName = realvalue2;
				colRealValue2.Caption = realvalue2;
			}
			else
				colRealValue2.Visible = false;
			if (count > 11)
			{
				string realvalue3 = Lib.ToString(dt.Columns[11]);
				colRealValue3.FieldName = realvalue3;
				colRealValue3.Caption = realvalue3;
			}
			else
				colRealValue3.Visible = false;
			if (count > 12)
			{
				string realvalue4 = Lib.ToString(dt.Columns[12]);
				colRealValue4.FieldName = realvalue4;
				colRealValue4.Caption = realvalue4;
			}
			else
				colRealValue4.Visible = false;
			if (count > 13)
			{
				string realvalue5 = Lib.ToString(dt.Columns[13]);
				colRealValue5.FieldName = realvalue5;
				colRealValue5.Caption = realvalue5;
			}
			else
				colRealValue5.Visible = false;
			grdData.DataSource = dt;
		}
		private void btnExportExecl_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount <= 0) return;
			string filePath = TextUtils.ExportExcelReturnFilePath(grvData, _QRCode);
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
				string[] arr = _QRCode.Split(' ');
				//	string[] arrOrder = arr[0].Split('-');

				workSheet.Cells[1, 1] = "Order No:";
				workSheet.Cells[1, 3] = _OrderCode;
				workSheet.Cells[2, 1] = "PID No:";
				workSheet.Cells[2, 3] = arr[1];

				ProductModel p = (ProductModel)ProductBO.Instance.FindByAttribute("ProductCode", arr[1].Trim())[0];

				workSheet.Cells[3, 1] = "Mô Tả";
				workSheet.Cells[3, 3] = p.ProductName;
				workSheet.Cells[4, 1] = "Ngày lắp:";
				workSheet.Cells[4, 3] = "'" + _DateLR.ToString("dd/MM/yyyy");
				//workSheet.Cells[4, 4] = "Mỡ sử dụng:";
				//workSheet.Cells[4, 5] = p.LoaiMo;
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


	}
}
