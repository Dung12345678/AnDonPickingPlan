using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IE.Utils;
using IE.Model;
using IE.Business;
using System.Collections;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System.Reflection;
using System.Threading;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using System.Diagnostics;
using DevExpress.XtraGrid.Localization;
using iTextSharp;
using iTextSharp.text.pdf;
using DevExpress.Utils;
using IOEx;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMS
{
	public partial class frmHistoryCheckH : _Forms
	{
		private string _pathFileConfigUpdate = Path.Combine(Application.StartupPath, "Download.txt");
		private string _pathFileConfigUpdateImage = Path.Combine(Application.StartupPath, "DownloadImage.txt");
		private string _pathFolderLocal;
		private string _pathUpdateServer;
		private string _pathFileVersion;
		string _pathError = Path.Combine(Application.StartupPath, "Errors");

		public string _QRCode;
		public string _OrderCode;
		public frmHistoryCheckH()
		{
			InitializeComponent();
		}

		private void frmHistoryCheck_Load(object sender, EventArgs e)
		{
			dtpFromDate.Value = DateTime.Now;
			dtpEndDate.Value = DateTime.Now;
		}

		void loadData()
		{
			DateTime from = new DateTime(dtpFromDate.Value.Year, dtpFromDate.Value.Month, dtpFromDate.Value.Day, 0, 0, 0);
			DateTime to = new DateTime(dtpEndDate.Value.Year, dtpEndDate.Value.Month, dtpEndDate.Value.Day, 23, 59, 59);

			DataTable dt = LibIE.LoadDataFromSP("spGetHistoryCheckData", "A"
				, new string[] { "@FromDate", "@EndDate", "@TextFind" }
				, new object[] { from, to, txtFindText.Text.Trim() });
			grdData.DataSource = dt;
		}
		private void btnFind_Click(object sender, EventArgs e)
		{
			loadData();
		}

		private void txtFindText_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				loadData();
			}
		}

		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0)
			{
				frmHistoryCheckDetailH frm = new frmHistoryCheckDetailH();
				frm.QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				frm.DateLR = TextUtils.ToDate3(grvData.GetFocusedRowCellValue(colDateLR));
				frm.OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				frm.Show();
			}
		}

		private void btnExportExecl_Click(object sender, EventArgs e)
		{
			if (grvData.FocusedRowHandle < 0)
			{
				return;
			}

			string path = "";

			FolderBrowserDialog f = new FolderBrowserDialog();
			if (f.ShowDialog() == DialogResult.OK)
			{
				path = f.SelectedPath;
			}
			else
			{
				return;
			}

			string _pPath = Application.StartupPath + "\\ResultCheckTemplate.xls";
			string orderCode = Lib.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
			string productCode = Lib.ToString(grvData.GetFocusedRowCellValue(colPID));

			using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang tạo file kết quả kiểm tra!"))
			{
				DataSet ds = LibIE.GetListDataFromSP("spGetProductWorkingInfo_ByOrder", new string[] { "@OrderCode", "@ProductCode" }, new object[] { orderCode, productCode });
				//   if (ds.Tables.Count <= 1) return;
				DataTable dtProduct = ds.Tables[0];
				if (ds.Tables.Count <= 1) return;
				DataTable dt = ds.Tables[1];

				DataRow r1 = dt.NewRow();
				dt.Rows.InsertAt(r1, 0);

				int count = dt.Rows.Count;

				Excel.Application app = default(Excel.Application);
				Excel.Workbook workBoook = default(Excel.Workbook);
				Excel.Worksheet workSheet = default(Excel.Worksheet);


				string fileNameHSTK = "KetQuaKiemTra_" + orderCode + "_" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xls";
				try
				{
					File.Copy(_pPath, path + "\\" + fileNameHSTK, true);
					app = new Excel.Application();
					app.Workbooks.Open(path + "\\" + fileNameHSTK);
					workBoook = app.Workbooks[1];
					workSheet = (Excel.Worksheet)workBoook.Worksheets[1];
					app.DisplayAlerts = false;

					string congDoan = "";
					int position = 0;

					//workSheet.Cells[2, 3] = "";
					workSheet.Cells[3, 2] = Lib.ToString(dtProduct.Rows[0]["ProductName"]);//tên
					workSheet.Cells[5, 6] = Lib.ToString(dtProduct.Rows[0]["MotorCode"]);//mã motor
					workSheet.Cells[6, 5] = Lib.ToString(dtProduct.Rows[0]["LoaiMo"]);//loại mỡ
					workSheet.Cells[6, 8] = Lib.ToString(dtProduct.Rows[0]["LuongMo"]);//lượng mỡ
					workSheet.Cells[2, 7] = grvData.GetFocusedRowCellDisplayText(colDateLR);//ngày lắp ráp
					workSheet.Cells[7, 7] = orderCode + "-1 " + productCode;

					workSheet.Cells[9, 10] = dt.Columns.Contains("1") ? orderCode + "-1" : "";
					workSheet.Cells[9, 11] = dt.Columns.Contains("2") ? orderCode + "-2" : "";
					workSheet.Cells[9, 12] = dt.Columns.Contains("3") ? orderCode + "-3" : "";
					workSheet.Cells[9, 13] = dt.Columns.Contains("4") ? orderCode + "-4" : "";
					workSheet.Cells[9, 14] = dt.Columns.Contains("5") ? orderCode + "-5" : "";
					workSheet.Cells[9, 15] = dt.Columns.Contains("6") ? orderCode + "-6" : "";

					for (int i = count - 1; i >= 0; i--)
					{
						DataRow r = dt.Rows[i];
						string maCongDoan = Lib.ToString(r["ProductStepCode"]);

						if (congDoan != maCongDoan)
						{
							((Excel.Range)workSheet.Rows[11]).Insert();
							((Excel.Range)workSheet.Rows[11]).Insert();
							workSheet.get_Range("A10", "Q10").Copy(workSheet.get_Range("A11", "Q11"));

							if (position > 1)
							{
								workSheet.get_Range("A14", "A" + (13 + position)).Merge(false);
								workSheet.get_Range("B14", "B" + (13 + position)).Merge(false);
							}
							position = 0;

							congDoan = maCongDoan;
							workSheet.Cells[12, 1] = maCongDoan;
							workSheet.Cells[12, 2] = "";
							workSheet.Cells[12, 3] = "";
							workSheet.Cells[12, 8] = "";
							workSheet.Cells[12, 9] = "Đánh giá";
							workSheet.Cells[12, 10] = dt.Columns.Contains("StatusResult1") ? Lib.ToString(r["StatusResult1"]) : "";
							workSheet.Cells[12, 11] = dt.Columns.Contains("StatusResult2") ? Lib.ToString(r["StatusResult2"]) : "";
							workSheet.Cells[12, 12] = dt.Columns.Contains("StatusResult3") ? Lib.ToString(r["StatusResult3"]) : "";
							workSheet.Cells[12, 13] = dt.Columns.Contains("StatusResult4") ? Lib.ToString(r["StatusResult4"]) : "";
							workSheet.Cells[12, 14] = dt.Columns.Contains("StatusResult5") ? Lib.ToString(r["StatusResult5"]) : "";
							workSheet.Cells[12, 15] = dt.Columns.Contains("StatusResult6") ? Lib.ToString(r["StatusResult6"]) : "";

							workSheet.get_Range("C12", "I12").Merge(false);
							workSheet.get_Range("A12", "Q12").Font.Size = 15;
							workSheet.get_Range("A12", "Q12").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

							position++;
						}
						((Excel.Range)workSheet.Rows[11]).Insert();
						workSheet.get_Range("A10", "Q10").Copy(workSheet.get_Range("A11", "Q11"));

						workSheet.Cells[12, 1] = maCongDoan;
						workSheet.Cells[12, 2] = Lib.ToString(r["ProductStepName"]);
						workSheet.Cells[12, 3] = Lib.ToString(r["WorkingName"]);
						workSheet.Cells[12, 8] = Lib.ToString(r["PeriodValue"]);
						workSheet.Cells[12, 9] = Lib.ToString(r["ValueTypeName"]);
						workSheet.Cells[12, 10] = dt.Columns.Contains("1") ? Lib.ToString(r["1"]) : "";
						workSheet.Cells[12, 11] = dt.Columns.Contains("2") ? Lib.ToString(r["2"]) : "";
						workSheet.Cells[12, 12] = dt.Columns.Contains("3") ? Lib.ToString(r["3"]) : "";
						workSheet.Cells[12, 13] = dt.Columns.Contains("4") ? Lib.ToString(r["4"]) : "";
						workSheet.Cells[12, 14] = dt.Columns.Contains("5") ? Lib.ToString(r["5"]) : "";
						workSheet.Cells[12, 15] = dt.Columns.Contains("6") ? Lib.ToString(r["6"]) : "";

						position++;
					}

					for (int i = 0; i < 5; i++)
					{
						((Excel.Range)workSheet.Rows[10]).Delete();
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
					Process.Start(path + "\\" + fileNameHSTK);
				}
			}
		}

		private void btnShowWorking_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0)
			{
				frmHistoryCheckOrderDetailH frm = new frmHistoryCheckOrderDetailH();
				frm._QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				frm._DateLR = TextUtils.ToDate3(grvData.GetFocusedRowCellValue(colDateLR));
				frm._OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				frm.Show();
			}
		}

		private void nGHEÂMTHANHToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (!File.Exists(_pathFileConfigUpdate)) return;
			try
			{
				_QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				_OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				string[] lines = File.ReadAllLines(_pathFileConfigUpdate);
				if (lines == null) return;
				if (lines.Length < 2) return;

				string[] stringSeparators = new string[] { "||" };
				string[] arr = lines[1].Split(stringSeparators, 4, StringSplitOptions.RemoveEmptyEntries);

				if (arr == null) return;
				if (arr.Length < 4) return;

				_pathFolderLocal = Path.Combine(Application.StartupPath, arr[1].Trim());
				_pathUpdateServer = arr[2].Trim();
				if (!Directory.Exists(_pathError))
				{
					Directory.CreateDirectory(_pathError);
				}
				if (!Directory.Exists(_pathFolderLocal))
				{
					Directory.CreateDirectory(_pathFolderLocal);
				}
				// File lưu Sound
				string path = "";
				FolderBrowserDialog f = new FolderBrowserDialog();
				if (f.ShowDialog() == DialogResult.OK)
				{
					path = f.SelectedPath;
				}
				else
				{
					path = _pathFolderLocal;
				}
				string[] arrr = _QRCode.Split(' ');
				string qrCode = arrr[0];
				string fileName = $"{_OrderCode}_{qrCode}_fwd.wav";
				if (DocUtils.CheckExits(_pathUpdateServer + $"\\{fileName}") == false)
				{
					MessageBox.Show("Không tìm thấy file âm thanh", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
				DocUtils.DownloadFile(path, fileName, Path.Combine(_pathUpdateServer, fileName));

				Process.Start($"{path}\\{fileName}");

			}
			catch { }
		}

		private void nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (!File.Exists(_pathFileConfigUpdate)) return;
			try
			{
				_QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				_OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				string[] lines = File.ReadAllLines(_pathFileConfigUpdate);
				if (lines == null) return;
				if (lines.Length < 2) return;

				string[] stringSeparators = new string[] { "||" };
				string[] arr = lines[1].Split(stringSeparators, 4, StringSplitOptions.RemoveEmptyEntries);

				if (arr == null) return;
				if (arr.Length < 4) return;

				_pathFolderLocal = Path.Combine(Application.StartupPath, arr[1].Trim());
				_pathUpdateServer = arr[2].Trim();
				if (!Directory.Exists(_pathError))
				{
					Directory.CreateDirectory(_pathError);
				}
				if (!Directory.Exists(_pathFolderLocal))
				{
					Directory.CreateDirectory(_pathFolderLocal);
				}
				// File lưu Sound
				string path = "";
				FolderBrowserDialog f = new FolderBrowserDialog();
				if (f.ShowDialog() == DialogResult.OK)
				{
					path = f.SelectedPath;
				}
				else
				{
					path = _pathFolderLocal;
				}
				string[] arrr = _QRCode.Split(' ');
				string qrCode = arrr[0];
				string fileName = $"{_OrderCode}_{qrCode}_rwd.wav";
				if (DocUtils.CheckExits(_pathUpdateServer + $"\\{fileName}") == false)
				{
					MessageBox.Show("Không tìm thấy file âm thanh", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
				DocUtils.DownloadFile(path, fileName, Path.Combine(_pathUpdateServer, fileName));

				Process.Start($"{path}\\{fileName}");

			}
			catch { }
		}

		private void xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (!File.Exists(_pathFileConfigUpdateImage)) return;
			try
			{
				_QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				_OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				string[] lines = File.ReadAllLines(_pathFileConfigUpdateImage);
				if (lines == null) return;
				if (lines.Length < 2) return;

				string[] stringSeparators = new string[] { "||" };
				string[] arr = lines[1].Split(stringSeparators, 4, StringSplitOptions.RemoveEmptyEntries);

				if (arr == null) return;
				if (arr.Length < 4) return;

				_pathFolderLocal = Path.Combine(Application.StartupPath, arr[1].Trim());
				_pathUpdateServer = arr[2].Trim();
				if (!Directory.Exists(_pathError))
				{
					Directory.CreateDirectory(_pathError);
				}
				if (!Directory.Exists(_pathFolderLocal))
				{
					Directory.CreateDirectory(_pathFolderLocal);
				}
				// File lưu Sound
				string path = "";
				FolderBrowserDialog f = new FolderBrowserDialog();
				if (f.ShowDialog() == DialogResult.OK)
				{
					path = f.SelectedPath;
				}
				else
				{
					path = _pathFolderLocal;
				}
				//string[] arrr = _QRCode.Split(' ');
				//string qrCode = arrr[0];
				string fileName = $"{_OrderCode}_{_QRCode}.jpg";
				DocUtils.DownloadFile(path, fileName, Path.Combine(_pathUpdateServer, fileName));

				Process.Start($"{path}\\{fileName}");

			}
			catch
			{
				MessageBox.Show("Lỗi đường dẫn không tìm thấy ảnh", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}
		/// <summary>
		/// Nghe lại file âm thanh của PCA theo số lần
		/// </summary>
		/// <param name="Number">số lần chạy</param>
		/// <param name="Rotation">Chiều quay F hoặc R</param>
		void RunSound(string Number,string Rotation)
		{
			if (!File.Exists(_pathFileConfigUpdate)) return;
			try
			{
				_QRCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colQRCode));
				_OrderCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colOrderCode));
				string[] lines = File.ReadAllLines(_pathFileConfigUpdate);
				if (lines == null) return;
				if (lines.Length < 2) return;

				string[] stringSeparators = new string[] { "||" };
				string[] arr = lines[1].Split(stringSeparators, 4, StringSplitOptions.RemoveEmptyEntries);

				if (arr == null) return;
				if (arr.Length < 4) return;

				_pathFolderLocal = Path.Combine(Application.StartupPath, arr[1].Trim());
				_pathUpdateServer = arr[2].Trim();
				if (!Directory.Exists(_pathError))
				{
					Directory.CreateDirectory(_pathError);
				}
				if (!Directory.Exists(_pathFolderLocal))
				{
					Directory.CreateDirectory(_pathFolderLocal);
				}
				// File lưu Sound
				string path = "";
				FolderBrowserDialog f = new FolderBrowserDialog();
				if (f.ShowDialog() == DialogResult.OK)
				{
					path = f.SelectedPath;
				}
				else
				{
					path = _pathFolderLocal;
				}
				string[] arrr = _QRCode.Split(' ');
				string qrCode = arrr[0];
				string fileName = "";
				if(Rotation=="F")
				{
					fileName = $"{_OrderCode}_{qrCode}_fwd" + Number + ".wav";
				}
				else if(Rotation == "R")
				{
					fileName = $"{_OrderCode}_{qrCode}_rwd" + Number + ".wav";
				}
				if (DocUtils.CheckExits(_pathUpdateServer + $"\\{fileName}") == false)
				{
					MessageBox.Show("Không tìm thấy file âm thanh", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
				DocUtils.DownloadFile(path, fileName, Path.Combine(_pathUpdateServer, fileName));

				Process.Start($"{path}\\{fileName}");

			}
			catch { }
		}

		private void nGHEÂMTHANHCHIỀUFToolStripMenuItem_Click(object sender, EventArgs e)
		{
			RunSound("","F");
		}

		private void nGHEÂMTHANHCHIỀURToolStripMenuItem_Click(object sender, EventArgs e)
		{
			RunSound("", "R");
		}

		private void nGHEÂMTHANHCHIỀUFToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			RunSound("1", "F");
		}

		private void nGHEÂMTHANHCHIỀURToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			RunSound("1", "R");
		}

		private void nGHEÂMTHANHCHIỀUFToolStripMenuItem2_Click(object sender, EventArgs e)
		{
			RunSound("2", "F");
		}

		private void nGHEÂMTHANHCHIỀURToolStripMenuItem2_Click(object sender, EventArgs e)
		{
			RunSound("2", "R");
		}

		private void nGHEÂMTHANHCHIỀUFToolStripMenuItem3_Click(object sender, EventArgs e)
		{
			RunSound("3", "F");
		}

		private void nGHEÂMTHANHCHIỀURToolStripMenuItem3_Click(object sender, EventArgs e)
		{
			RunSound("3", "R");
		}
	}
}
