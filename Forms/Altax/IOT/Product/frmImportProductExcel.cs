using BMS.Business;
using BMS.Model;
using ExcelDataReader;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmImportProductExcel : _Forms
	{
		public frmImportProductExcel()
		{
			InitializeComponent();
		}
		private DataSet ds;
		string getSql(string periodValue, string min, string max, string sortOrder, string stepCode, int productID)
		{
			string sql = "";
			sql = string.Format(@"UPDATE dbo.ProductWorking SET  PeriodValue = N'{0}', MinValue = {1}, MaxValue = {2} 
                                    WHERE  ProductStepCode = '{3}' AND SortOrder IN ({4})  AND ProductID = {5}"
								, periodValue, min, max, stepCode, sortOrder, productID);
			return sql;
		}

		string getSqlGroup(string periodValue, string min, string max, string sortOrder, string stepCode, int productID, string groupCode)
		{
			string sql = "";
			sql = string.Format(@"UPDATE w 
                                    SET w.PeriodValue = N'{0}', w.MinValue = {1}, w.MaxValue = {2} 
                                    FROM dbo.ProductWorking w 
                                        JOIN dbo.Product p ON p.ID = w.ProductID AND p.ProductGroupCode = '{3}' AND p.ID = {4}
                                    WHERE w.ProductStepCode ='{5}' AND w.SortOrder IN ({6})"
								, periodValue, min, max, groupCode, productID, stepCode, sortOrder);
			return sql;
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
				progressBar1.Maximum = grvData.RowCount - 3;
				txtRate.Text = "";
				start = DateTime.Now;
				enableControl(false);
				backgroundWorker1.RunWorkerAsync();
			}
		}

		private void btnBrowse_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
		{
			grvData.Columns.Clear();
			OpenFileDialog ofd = new OpenFileDialog();
			if (chkAutoCheck.Checked)
			{
				OpenFileDialog openFileDialog1 = new OpenFileDialog();
				var result = openFileDialog1.ShowDialog();
				if (result == DialogResult.OK)
				{
					btnBrowse.Text = openFileDialog1.FileName;
				}
				else if (result == DialogResult.Cancel)
				{
					return;
				}

				try
				{
					var stream = new FileStream(btnBrowse.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

					var sw = new Stopwatch();
					sw.Start();

					IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

					var openTiming = sw.ElapsedMilliseconds;

					ds = reader.AsDataSet(new ExcelDataSetConfiguration()
					{
						UseColumnDataType = false,
						ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
						{
							UseHeaderRow = false
						}
					});
					//toolStripStatusLabel1.Text = "Elapsed: " + sw.ElapsedMilliseconds.ToString() + " ms (" + openTiming.ToString() + " ms to open)";

					var tablenames = GetTablenames(ds.Tables);

					cboSheet.DataSource = tablenames;

					if (tablenames.Count > 0)
						cboSheet.SelectedIndex = 0;
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
				}

			}
			else
			{
				if (ofd.ShowDialog() == DialogResult.OK)
				{
					btnBrowse.Text = ofd.FileName;
					cboSheet.DataSource = null;
					cboSheet.DataSource = TextUtils.ListSheetInExcel(ofd.FileName);
				}
			}
		}
		private static IList<string> GetTablenames(DataTableCollection tables)
		{
			var tableList = new List<string>();
			foreach (var table in tables)
			{
				tableList.Add(table.ToString());
			}

			return tableList;
		}
		private void cboSheet_SelectionChangeCommitted(object sender, EventArgs e)
		{
			grvData.Columns.Clear();
			if (chkAutoCheck.Checked)
			{
				try
				{
					var tablename = cboSheet.SelectedItem.ToString();

					grdData.DataSource = ds; // dataset
					grdData.DataMember = tablename;
				}
				catch (Exception ex)
				{
					TextUtils.ShowError(ex);
					grdData.DataSource = null;
				}
			}
			else
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
		}

		private void toolStripButton1_Click(object sender, EventArgs e)
		{
			TextUtils.ExportExcel(grvData);
		}

		private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			//txtRate.Text = "";

			DataTable dtGroup = Lib.Select("SELECT ID,ProductGroupCode FROM dbo.ProductGroup");
			List<BaseModel> lstProduct = ProductBO.Instance.GetAllListObject();
			//object[] lst = arrProduct.ToArray();
			//progressBar1.Minimum = 0;

			int rowCount = grvData.RowCount;

			List<string> lstError = new List<string>();
			char spt = '~';
			for (int i = 0; i < rowCount; i++)
			{
				try
				{
					if (i < 3) continue;
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i - 2; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i - 2, rowCount - 3); }));
					string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
					string groupCode = Lib.ToString(grvData.GetRowCellValue(i, "F2"));
					//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
					if (string.IsNullOrEmpty(groupCode) || string.IsNullOrEmpty(productCode))
					{
						continue;
					}

					DataRow[] arrG = dtGroup.Select(string.Format("ProductGroupCode = '{0}'", groupCode));
					int groupID = 0;
					if (arrG.Length > 0)
					{
						groupID = Lib.ToInt(arrG[0]["ID"]);
					}
					else
					{
						continue;
					}

					ProductModel productModel = new ProductModel();
					var lstObject = lstProduct.FirstOrDefault(o => ((ProductModel)o).ProductCode.ToLower().Trim() == productCode.ToLower().Trim());
					//ArrayList arrP = ProductBO.Instance.FindByAttribute("ProductCode", productCode);
					if (lstObject != null)
					{
						productModel = (ProductModel)lstObject;
					}

					#region SetValue
					productModel.ProductGroupID = groupID;
					productModel.ProductGroupCode = groupCode;
					productModel.ProductCode = productCode;
					productModel.ProductName = Lib.ToString(grvData.GetRowCellValue(i, "F3"));//Mô tả-describe
					productModel.ProductTypeCode = Lib.ToString(grvData.GetRowCellValue(i, "F4"));//Mô tả-describe

					productModel.RatioCode = Lib.ToString(grvData.GetRowCellValue(i, "F5"));//Ratio
					productModel.PinNumber = Lib.ToString(grvData.GetRowCellValue(i, "F6"));//SỐ PIN NGOÀI
					productModel.PinNumberIN = Lib.ToString(grvData.GetRowCellValue(i, "F7"));//SỐ PIN TRONG

					string valueString = Lib.ToString(grvData.GetRowCellValue(i, "F8"));

					productModel.SizeSSSMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//KÍCH THƯỚC CỤM TRỤC SSS
					productModel.SizeSSSMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//KÍCH THƯỚC CỤM TRỤC SSS
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F9"));
					productModel.DoDao1Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 1
					productModel.DoDao1Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 1
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F10"));
					productModel.DoDao2Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 2
					productModel.DoDao2Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 2
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F11"));
					productModel.DoDao3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 3
					productModel.DoDao3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 3
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F12"));
					productModel.DoDao4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 4
					productModel.DoDao4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 4? Khe hở cho phép

					productModel.MotorCode = Lib.ToString(grvData.GetRowCellValue(i, "F13"));//MÃ MOTOR
					productModel.LoaiMo = Lib.ToString(grvData.GetRowCellValue(i, "F14"));//LOẠI MỠ
					productModel.LuongMo = Lib.ToDecimal(grvData.GetRowCellValue(i, "F15"));//LƯỢNG MỠ TỔNG

					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F16"));
					productModel.LuongMo1Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho trục tốc độ thấp
					productModel.LuongMo1Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho trục tốc độ thấp
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F17"));
					productModel.LuongMo2Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lượng mỡ cho hộp
					productModel.LuongMo2Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lượng mỡ cho hộp
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F18"));
					productModel.LuongMo3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho vòng bi chịu tải & Carrier SSS
					productModel.LuongMo3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho vòng bi chịu tải & Carrier SSS
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F19"));
					productModel.LuongMo4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//LƯƠNG MỠ CHO VÒNG BI KHÔNG TẢI
					productModel.LuongMo4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//LƯƠNG MỠ CHO VÒNG BI KHÔNG TẢI
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F20"));
					productModel.LuongMo5Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho Carrier
					productModel.LuongMo5Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho Carrier
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F21"));
					productModel.LuongMo6Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512
					productModel.LuongMo6Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512

					productModel.MEDEDAY = Lib.ToDate2(grvData.GetRowCellValue(i, "F22"));//DATE
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F23"));
					productModel.LucCheckGearMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check gear Unit
					productModel.LucCheckGearMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check gear Unit

					productModel.GunNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F24"));//Súng
					productModel.JobNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F25"));//JOB
					productModel.BuLongNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F26"));//SỐ LUỢNG BULONG
					productModel.ShootNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F27"));//SỐ LẦN BẮN
					productModel.QtyOcBanGa = Lib.ToInt(grvData.GetRowCellValue(i, "F28"));//BẮN GÁ
					productModel.QtyOcBanThat = productModel.SetForce = Lib.ToInt(grvData.GetRowCellValue(i, "F29"));//ĐẶT LỰC

					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F30"));
					productModel.LucCheckGearMotorMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check Gear Motor
					productModel.LucCheckGearMotorMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check Gear Motor

					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F31"));
					productModel.LucCheck3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check cụm trục tốc độ thấp 511-512
					productModel.LucCheck3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check cụm trục tốc độ thấp 511-512
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F32"));
					productModel.LucCheck4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check phần xuất lực Naknishi
					productModel.LucCheck4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check phần xuất lực Naknishi
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F33"));
					productModel.LucCheck5Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check Ốc chặn mỡ Nakanishi, 6000
					productModel.LucCheck5Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check Ốc chặn mỡ Nakanishi, 6000

					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F34"));
					productModel.SpaceCenterMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//KHOẢNG CÁCH TRỤC TRUNG TÂM.
					productModel.SpaceCenterMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//KHOẢNG CÁCH TRỤC TRUNG TÂM.

					productModel.MURATA = Lib.ToString(grvData.GetRowCellValue(i, "F35"));//MURATA
					productModel.CoilCode = Lib.ToString(grvData.GetRowCellValue(i, "F36"));//MÃ COIL
					productModel.DienAp = Lib.ToDecimal(grvData.GetRowCellValue(i, "F37"));//Điện áp (V)
					productModel.TanSo = Lib.ToDecimal(grvData.GetRowCellValue(i, "F38"));//Tần số

					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F39"));
					productModel.DongDienMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Dòng điện Io
					productModel.DongDienMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Dòng điện Io
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F40"));
					productModel.NhapLucMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Nhập lực Wo
					productModel.NhapLucMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Nhập lực Wo
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F41"));
					productModel.VongQuayMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Vòng quay
					productModel.VongQuayMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Vòng quay
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F42"));
					productModel.TiengOnMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//TIẾNG ỒN
					productModel.TiengOnMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//TIẾNG ỒN
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F43"));
					productModel.DoRungFFTMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ RUNG FFT
					productModel.DoRungFFTMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ RUNG FFT
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F44"));
					productModel.SpaceCenterMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//khoảng cách A giữa hai mặt vòng bi Cycloid disk
					productModel.SpaceCenterMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//khoảng cách A giữa hai mặt vòng bi Cycloid disk
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F45"));
					productModel.DifferenceMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Độ chênh lệch lớn nhất tại 3 vị trí đo 0.03mm
					productModel.DifferenceMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Độ chênh lệch lớn nhất tại 3 vị trí đo 0.03mm

					//Các mục kiểm tra cho công đoạn 1
					productModel.SSS = Lib.ToString(grvData.GetRowCellValue(i, "F46"));
					productModel.CARRIER = Lib.ToString(grvData.GetRowCellValue(i, "F47"));
					productModel.CYCLODISK = Lib.ToString(grvData.GetRowCellValue(i, "F48"));
					productModel.RGH = Lib.ToString(grvData.GetRowCellValue(i, "F49"));
					productModel.BEFORECOVER = Lib.ToString(grvData.GetRowCellValue(i, "F50"));
					productModel.AFTERCOVER = Lib.ToString(grvData.GetRowCellValue(i, "F51"));
					productModel.SSS2 = Lib.ToString(grvData.GetRowCellValue(i, "F52"));
					productModel.CARRIER2 = Lib.ToString(grvData.GetRowCellValue(i, "F53"));
					productModel.CYCLODISK2 = Lib.ToString(grvData.GetRowCellValue(i, "F54"));
					productModel.RGH2 = Lib.ToString(grvData.GetRowCellValue(i, "F55"));
					productModel.BEFORECOVER2 = Lib.ToString(grvData.GetRowCellValue(i, "F56"));
					productModel.AFTERCOVER2 = Lib.ToString(grvData.GetRowCellValue(i, "F57"));

					//nvthao thêm ngày 03.01.2020
					//Thêm trường Đích lưu cho sản phẩm
					productModel.Goal = Lib.ToString(grvData.GetRowCellValue(i, "F58"));


					#endregion

					List<string> lstCommand = new List<string>();
					if (!chkWorking.Checked)
					{
						if (productModel.ID == 0)
						{
							productModel.ID = (int)ProductBO.Instance.Insert(productModel);
							//Thêm các hạng mục check theo nhóm sản phẩm
							string sqlInitWorking = string.Format("EXEC dbo.CreateProductWorking_ByGroupID {0},{1}", groupID, productModel.ID);
							lstCommand.Add(sqlInitWorking);
							// TextUtils.ExcuteSQL(sqlInitWorking);
							lstProduct.Add(productModel);
						}
						else
						{
							ProductBO.Instance.Update(productModel);
						}
					}

					//string sqlReset = string.Format(@"UPDATE dbo.ProductWorking SET PeriodValue = '', MinValue = 0, MaxValue = 0 
					//                                    WHERE ProductID = {0}", productModel.ID);
					//lstCommand.Add(sqlReset);

					//Sinh các câu lệnh update động các mục kiểm tra
					for (int j = 0; j < grvData.Columns.Count; j++)
					{
						string rule = Lib.ToString(grvData.GetRowCellValue(3, grvData.Columns[j]));
						if (string.IsNullOrEmpty(rule)) continue;
						string value = Lib.ToString(grvData.GetRowCellValue(i, grvData.Columns[j]));

						if (string.IsNullOrEmpty(value)) value = "";

						string min = "";
						string max = "";
						if (value.Contains(spt))
						{
							string[] arrValue = value.Split(spt);
							min = arrValue[0].Trim();
							max = arrValue[1].Trim();
						}
						else
						{
							min = max = Lib.ToDecimal(value).ToString();
						}

						string[] arrCD = rule.Split(';');
						for (int h = 0; h < arrCD.Length; h++)
						{
							string[] arrW = arrCD[h].Trim().Split('_');
							if (arrW.Length <= 1) continue;

							if (arrW.Length == 3)//chứa cả nhóm sản phẩm,mã công đoạn, stt mục kiểm tra
							{
								string groupCode1 = arrW[0].Trim();
								if (groupCode.ToLower() != groupCode1.ToLower()) continue;

								string stepCode = arrW[1].Trim();
								string sortOrder = arrW[2].Trim();
								string sql = getSqlGroup(value, min, max, sortOrder, stepCode, productModel.ID, groupCode1);
								lstCommand.Add(sql);
							}

							if (arrW.Length == 2)//chỉ chứa mã công đoạn, stt mục kiểm tra
							{
								string stepCode = arrW[0].Trim();

								string sortOrder = arrW[1].Trim();
								string sql = getSql(value, min, max, sortOrder, stepCode, productModel.ID);
								lstCommand.Add(sql);
							}
						}
					}

					//string sqlUpdateWorking = string.Format("EXEC spUpdateProductWorking_ByProductID_New {0},{1}", productModel.ID, productModel.MURATA == "3" ? 0 : 1);
					//lstCommand.Add(sqlUpdateWorking);
					//TextUtils.ExcuteSQL(sqlUpdateWorking);

					if (lstCommand.Count > 0)
					{
						Lib.ExcuteSQLImport(string.Join(";", lstCommand));
					}
				}
				catch (Exception ex)
				{
					//lstError.Add(productCode);
					MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i + Environment.NewLine + ex.ToString());
				}
				//stopwatch.Stop();
				//MessageBox.Show(stopwatch.ElapsedMilliseconds.ToString());
			}

		}

		private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
		{

		}

		private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show(start.ToString() + " - " + DateTime.Now.ToString());
			enableControl(true);
		}

		private void frmImportProductExcel_Load(object sender, EventArgs e)
		{

		}
	}
}
