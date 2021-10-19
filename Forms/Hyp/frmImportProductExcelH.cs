using ExcelDataReader;
using IE.Business;
using IE.Model;
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
	public partial class frmImportProductExcelH : Form
	{
		public frmImportProductExcelH()
		{
			InitializeComponent();
		}
		private DataSet ds;
		DataTable dt;
		string getSql(string periodValue, string min, string max, string sortOrder, string stepCode, int productID)
		{
			string sql = "";
			sql = string.Format(@"UPDATE dbo.ProductWorking SET  PeriodValue = N'{0}', MinValue = {1}, MaxValue = {2} 
			                        WHERE  ProductStepCode = '{3}' AND SortOrder IN ({4})  AND ProductID = {5}"
								, periodValue, min, max, stepCode, sortOrder, productID);

			//sql = string.Format($"Exec spUpdateProductWorkingNoGroupByImportExcel @periodValue=N'{periodValue}', @min={min}, @max={max}, @productID={productID}, @stepCode=N'{stepCode}', @sortOrder={sortOrder},@createdBy=N'{Global.AppUserName}'");
			return sql;
		}

		string getSqlGroup(string periodValue, string min, string max, string sortOrder, string stepCode, int productID, string groupCode)
		{
			string sql = "";
			sql = string.Format(@"UPDATE w 
                                    SET w.PeriodValue = '{0}', w.MinValue = {1}, w.MaxValue = {2} 
                                    FROM dbo.ProductWorking w 
                                        JOIN dbo.Product p ON p.ID = w.ProductID AND p.ProductGroupCode = '{3}' AND p.ID = {4}
                                    WHERE w.ProductStepCode ='{5}' AND w.SortOrder IN ({6})"
								, periodValue, min, max, groupCode, productID, stepCode, sortOrder);
			//sql = string.Format(@"Exec spUpdateProductWorkingByImportExcel @periodValue=N'{0}', @min={1}, @max={2}, @groupCode=N'{3}', @productID={4}, @stepCode=N'{5}', @sortOrder={6},@createdBy=N'{7}'",
			//             periodValue, min, max, groupCode, productID, stepCode, sortOrder, Global.AppUserName);

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
					DataTable dtt = ds.Tables[cboSheet.SelectedIndex];
					dtt.Columns.Add("DATEF6", typeof(DateTime));
					for (int i = 0; i < dtt.Rows.Count; i++)
					{
						if (i < 3) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F2"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F6"]).Trim() != "")
							dtt.Rows[i]["DATEF6"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F6"]));
					}

					var tablename = cboSheet.SelectedItem.ToString();

					grdData.DataSource = dtt;
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
			else
			{
				try
				{
					DataTable dt = TextUtils.ExcelToDatatableNoHeader(btnBrowse.Text, cboSheet.SelectedValue.ToString());

					// thêm 1 cột date mới
					dt.Columns.Add("DATEF6", typeof(DateTime));
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						if (i < 3) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F2"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F6"]).Trim() != "")
							dt.Rows[i]["DATEF6"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F6"]));
					}

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

			DataTable dtGroup = LibIE.Select("SELECT ID,ProductGroupCode FROM dbo.ProductGroup");
			List<BaseModel> lstProduct = ProductBO.Instance.GetAllListObject();

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
					string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));// productModel ID
					string groupCode = Lib.ToString(grvData.GetRowCellValue(i, "F2")); // Nhóm sản phẩm
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
				//	var lstObject = lstProduct.FirstOrDefault(o => ((ProductModel)o).ProductCode.Trim().ToLower() == productCode.Trim().ToLower());
					ArrayList arrP = ProductBO.Instance.FindByAttribute("ProductCode", productCode);
					int OldGroupID = 0;
					if (arrP.Count>0)
					{
						productModel = (ProductModel)arrP[0];
						OldGroupID = productModel.ProductGroupID;
						if (OldGroupID != groupID)
						{
							ProductWorkingBO.Instance.DeleteByAttribute("ProductID", productModel.ID);
							ProductStepBO.Instance.DeleteByAttribute("ProductID", productModel.ID);
							ProductBO.Instance.Delete(productModel.ID);
							productModel.ID = 0;
						}
					}

					//if (lstObject != null)
					//{
					//	productModel = (ProductModel)lstObject;
					//	OldGroupID = productModel.ProductGroupID;
					//	if (OldGroupID != groupID)
					//	{
					//		ProductWorkingBO.Instance.DeleteByAttribute("ProductID", productModel.ID);
					//		ProductStepBO.Instance.DeleteByAttribute("ProductID", productModel.ID);
					//		ProductBO.Instance.Delete(productModel.ID);
					//		productModel.ID = 0;
					//	}
					//}

					#region SetValue
					string valueString = "";
					productModel.ProductGroupID = groupID;
					productModel.ProductGroupCode = groupCode;
					productModel.ProductCode = productCode;
					productModel.ProductName = Lib.ToString(grvData.GetRowCellValue(i, "F3"));//Mô tả-sản phẩm
					productModel.ProductTypeCode = Lib.ToString(grvData.GetRowCellValue(i, "F4"));//Loại sản phẩm  
					productModel.RatioCode = Lib.ToString(grvData.GetRowCellValue(i, "F5"));//Ratio giảm tốc

					productModel.ProductDate = Lib.ToDate3(grvData.GetRowCellValue(i, "F6"));//ngày lập				
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F6")) == new DateTime(1950, 1, 1))
					{
						productModel.ProductDate = Lib.ToDate2(grvData.GetRowCellValue(i, "DATEF6").ToString());
					}

					productModel.XuatLucSealCap = Lib.ToString(grvData.GetRowCellValue(i, "F7"));//Hướng trục xuất lực
					productModel.LoaiMoNapDau = Lib.ToString(grvData.GetRowCellValue(i, "F8"));// loại mỡ nắp dầu   

					productModel.LoaiMoBomHopSo = Lib.ToString(grvData.GetRowCellValue(i, "F10"));// loại mỡ bơm hộp số  

					productModel.LucCheckGear = valueString = Lib.ToString(grvData.GetRowCellValue(i, "F37"));
					productModel.LucCheckGearMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check gear Unit
					productModel.LucCheckGearMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check gear Unit

					productModel.GunNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F38"));//Súng
					productModel.JobNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F39"));//JOB
					productModel.BuLongNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F40"));//SỐ LUỢNG BULONG
					productModel.ShootNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F41"));//SỐ LẦN BẮN
					productModel.QtyOcBanGa = Lib.ToInt(grvData.GetRowCellValue(i, "F42"));//BẮN GÁ
					productModel.QtyOcBanThat = productModel.SetForce = Lib.ToInt(grvData.GetRowCellValue(i, "F43"));//ĐẶT LỰC



					productModel.MotorCode = Lib.ToString(grvData.GetRowCellValue(i, "F53"));//MÃ MOTOR
					productModel.CoilCode = Lib.ToString(grvData.GetRowCellValue(i, "F56"));//MÃ COIL
					productModel.DienAp = Lib.ToDecimal(grvData.GetRowCellValue(i, "F57"));//Điện áp (V)
					productModel.TanSo = Lib.ToDecimal(grvData.GetRowCellValue(i, "F58"));//Tần số
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F59"));
					productModel.DongDienMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Dòng điện Io
					productModel.DongDienMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Dòng điện Io
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F60"));
					productModel.NhapLucMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Nhập lực Wo
					productModel.NhapLucMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Nhập lực Wo
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F61"));
					productModel.VongQuayMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Vòng quay
					productModel.VongQuayMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Vòng quay
					productModel.TiengOn = valueString = Lib.ToString(grvData.GetRowCellValue(i, "F62"));
					productModel.TiengOnMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//TIẾNG ỒN
					productModel.TiengOnMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//TIẾNG ỒN
					valueString = Lib.ToString(grvData.GetRowCellValue(i, "F63"));
					productModel.DoRungFFTMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ RUNG FFT
					productModel.DoRungFFTMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ RUNG FFT
					productModel.Goal = Lib.ToString(grvData.GetRowCellValue(i, "F64"));
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
							lstProduct.Add(productModel);
						}
						else
						{
							ProductBO.Instance.Update(productModel);
						}
					}
					//Sinh các câu lệnh update động các mục kiểm tra
					for (int j = 0; j < grvData.Columns.Count; j++)
					{
						string rule = Lib.ToString(grvData.GetRowCellValue(3, grvData.Columns[j]));
						if (string.IsNullOrEmpty(rule)) continue;
						string value = Lib.ToString(grvData.GetRowCellValue(i, grvData.Columns[j]));
						if (string.IsNullOrEmpty(value)) value = "";
						value = value.Replace("'", "");

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

					if (lstCommand.Count > 0)
					{
						string sqlRun = string.Join(";", lstCommand);
						LibIE.ExcuteSQL(sqlRun);
					}
				}
				catch (Exception ex)
				{
					//lstError.Add(productCode);
					MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i);
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

		private void chkAutoCheck_CheckedChanged(object sender, EventArgs e)
		{

		}

		private void frmImportProductExcelH_Load(object sender, EventArgs e)
		{

		}
	}
}
