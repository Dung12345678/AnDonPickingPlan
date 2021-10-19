using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Media;

namespace BMS
{
	public partial class frmHistoryCheckDetailH : _Forms
	{
		public int CheckHistoryID;
		public string QRCode;
		public DateTime DateLR;
		public string OrderCode;
		bool _isErrorExtract = false;

		private string _pathFileConfigUpdate = Path.Combine(Application.StartupPath, "Download.txt");
		string _extractPath = Path.GetFullPath(Application.StartupPath);
		private string _pathFolderLocal;
		private string _pathUpdateServer;
		private string _pathFileVersion;
		string _pathError = Path.Combine(Application.StartupPath, "Errors");

		public frmHistoryCheckDetailH()
		{
			InitializeComponent();
		}

		private void frmHistoryCheckDetail_Load(object sender, EventArgs e)
		{
			loadData();
			DocUtils.InitFTPQLSX();
			//updateVersion();

		}
		void loadData()
		{
			//spLoadCheckHistoryDetail//CheckHistoryID
			DataTable dt = LibIE.LoadDataFromSP("spGetHistoryCheckDataDetail", "A"
				, new string[] { "QRCode" }
				, new object[] { QRCode });
			grdData.DataSource = dt;
		}

		private void btnExportExecl_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount <= 0) return;
			string filePath = TextUtils.ExportExcelReturnFilePath(grvData, QRCode);
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
				string[] arr = QRCode.Split(' ');
				string[] arrOrder = arr[0].Split('-');

				workSheet.Cells[1, 1] = "Order No:";
				workSheet.Cells[1, 3] = arrOrder[0];
				workSheet.Cells[2, 1] = "PID No:";
				workSheet.Cells[2, 3] = arr[1];

				ProductModel p = (ProductModel)ProductBO.Instance.FindByAttribute("ProductCode", arr[1].Trim())[0];

				workSheet.Cells[3, 1] = "Mô Tả";
				workSheet.Cells[3, 3] = p.ProductName;
				workSheet.Cells[4, 1] = "Ngày lắp:";
				workSheet.Cells[4, 3] = "'" + DateLR.ToString("dd/MM/yyyy HH:mm:ss");
				workSheet.Cells[4, 4] = "Mỡ sử dụng:";
				workSheet.Cells[4, 5] = p.LoaiMo;
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
