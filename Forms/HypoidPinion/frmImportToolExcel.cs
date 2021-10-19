using ExcelDataReader;
using HP.Business;
using HP.Model;
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
	public partial class frmImportToolExcel : Form
	{
		public frmImportToolExcel()
		{
			InitializeComponent();
		}
		private DataSet ds;
		DataTable dt;
		string getSql(string periodValue, string min, string max, string sortOrder, string stepCode, int productID)
		{
			string sql = "";
			sql = string.Format(@"UPDATE dbo.KnifeJigWorking SET  PeriodValue = N'{0}', MinValue = {1}, MaxValue = {2}
			                        WHERE  ProductStepCode = '{3}' AND SortOrder IN ({4})  AND KnifeDetailListID = {5}"
								, periodValue, min, max, stepCode, sortOrder, productID);

			//sql = string.Format($"Exec spUpdateProductWorkingNoGroupByImportExcel @periodValue=N'{periodValue}', @min={min}, @max={max}, @productID={productID}, @stepCode=N'{stepCode}', @sortOrder={sortOrder},@createdBy=N'{Global.AppUserName}'");
			return sql;
		}

		string getSqlGroup(string periodValue, string min, string max, string sortOrder, string stepCode, int productID, string groupCode)
		{
			string sql = "";
			sql = string.Format(@"UPDATE w 
                                    SET w.PeriodValue = '{0}', w.MinValue = {1}, w.MaxValue = {2} 
                                    FROM dbo.KnifeJigWorking w 
                                        JOIN dbo.KnifeDetailList p ON p.ID = w.KnifeDetailListID AND p.ID = {4}
										JOIN dbo.KnifeGroup k ON k.ID = p.KnifeGroupID
                                    WHERE w.ProductStepCode ='{5}' AND w.SortOrder IN ({6}) AND k.KnifeGroupCode='{3}'"
								, periodValue, min, max, groupCode, productID, stepCode, sortOrder);
			//sql = string.Format(@"Exec spUpdateProductWorkingByImportExcel @periodValue=N'{0}', @min={1}, @max={2}, @groupCode=N'{3}', @productID={4}, @stepCode=N'{5}', @sortOrder={6},@createdBy=N'{7}'",
			//             periodValue, min, max, groupCode, productID, stepCode, sortOrder, Global.AppUserName);

			return sql;
		}

		string getSqlGroupFrequency(string Frequency, string sortOrder, string stepCode, int productID, string groupCode)
		{
			string sql = "";
			sql = string.Format(@"UPDATE w 
                                    SET w.Frequency='{0}'
                                    FROM dbo.KnifeJigWorking w 
                                        JOIN dbo.KnifeDetailList p ON p.ID = w.KnifeDetailListID AND p.ID = {2}
										JOIN dbo.KnifeGroup k ON k.ID = p.KnifeGroupID
                                    WHERE w.ProductStepCode ='{3}' AND w.SortOrder IN ({4}) AND k.KnifeGroupCode='{1}'"
								, Frequency, groupCode, productID, stepCode, sortOrder);
			//sql = string.Format(@"Exec spUpdateProductWorkingByImportExcel @periodValue=N'{0}', @min={1}, @max={2}, @groupCode=N'{3}', @productID={4}, @stepCode=N'{5}', @sortOrder={6},@createdBy=N'{7}'",
			//             periodValue, min, max, groupCode, productID, stepCode, sortOrder, Global.AppUserName);

			return sql;
		}
		string getSqlFrequency(string Frequency, string sortOrder, string stepCode, int productID)
		{
			string sql = "";
			sql = string.Format(@"UPDATE dbo.KnifeJigWorking SET  Frequency='{0}'
			                        WHERE  ProductStepCode = '{1}' AND SortOrder IN ({2})  AND KnifeDetailListID = {3}"
								, Frequency, stepCode, sortOrder, productID);

			//sql = string.Format($"Exec spUpdateProductWorkingNoGroupByImportExcel @periodValue=N'{periodValue}', @min={min}, @max={max}, @productID={productID}, @stepCode=N'{stepCode}', @sortOrder={sortOrder},@createdBy=N'{Global.AppUserName}'");
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
			DataTable dtGroup = new DataTable();
			this.Invoke((MethodInvoker)delegate
			{
				if (cboType.SelectedIndex == 0) return;

				dtGroup = TextUtilsHP.Select($"SELECT ID,KnifeGroupCode FROM [HypoidPinion].[dbo].[KnifeGroup] Where Type={cboType.SelectedIndex - 1}");

				int rowCount = grvData.RowCount;

				char spt = '~';
				char dau = '/';
				for (int i = 0; i < rowCount; i++)
				{
					try
					{
						if (i < 3) continue;
						progressBar1.Invoke((Action)(() => { progressBar1.Value = i - 2; }));
						txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i - 2, rowCount - 3); }));
						string KnifeCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));// KnifeCode ID
						string KnifeGroupCode = Lib.ToString(grvData.GetRowCellValue(i, "F2")); // Nhóm sản phẩm
																								//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
						if (string.IsNullOrEmpty(KnifeGroupCode) || string.IsNullOrEmpty(KnifeCode))
						{
							continue;
						}

						DataRow[] arrG = dtGroup.Select(string.Format("KnifeGroupCode = '{0}'", KnifeGroupCode));
						int groupID = 0;
						if (arrG.Length > 0)
						{
							groupID = Lib.ToInt(arrG[0]["ID"]);
						}
						else
						{
							continue;
						}
						//Thước
						if (cboType.SelectedIndex == 4)
						{
							ToolsModel toolsModel = new ToolsModel();
							ArrayList arrTool = ToolsBO.Instance.FindByAttribute("Code", KnifeCode);
							//var lstObject = lstProduct.FirstOrDefault(o => ((ProductModel)o).ProductCode.Trim().ToLower() == productCode.Trim().ToLower());
							int OldGroupID = 0;
							if (arrTool.Count > 0)
							{
								toolsModel = (ToolsModel)arrTool[0];
								OldGroupID = toolsModel.GroupID;
								if (OldGroupID != groupID)
								{
									//Xóa trường cũ đi và thêm trường mới vào
									KnifeJigWorkingBO.Instance.DeleteByAttribute("KnifeDetailListID", toolsModel.ID);
									KnifeStepBO.Instance.DeleteByAttribute("KnifeDetailListID", toolsModel.ID);
									ToolsBO.Instance.Delete(toolsModel.ID);
									toolsModel.ID = 0;
								}
							}

							//Insert dụng cụ đo

							toolsModel.GroupID = groupID;
							toolsModel.Code = KnifeCode;
							toolsModel.Name = Lib.ToString(grvData.GetRowCellValue(i, "F3"));//Tên dụng cụ
							toolsModel.Std = Lib.ToInt(grvData.GetRowCellValue(i, "F7"));//STD
							toolsModel.Min = Lib.ToDecimal(grvData.GetRowCellValue(i, "F9"));//Min
							toolsModel.Max = Lib.ToDecimal(grvData.GetRowCellValue(i, "F10"));//Max
							toolsModel.NameDisplay = toolsModel.Name + toolsModel.Min + toolsModel.Max;
							toolsModel.ConditionDisplay = toolsModel.Min + "-" + toolsModel.Max;
							toolsModel.DateOld = Lib.ToDate2(grvData.GetRowCellValue(i, "F11"));// khai báo ngày để hiện của lịch sử 
							toolsModel.FrequencyCheck = Lib.ToString(grvData.GetRowCellValue(i, "F12")); // tần suất kiểm tra theo ngày tháng năm
							if (toolsModel.FrequencyCheck.Trim().ToUpper() == "TUẦN" || toolsModel.FrequencyCheck.Trim().ToUpper() == "WEEK")
							{
								toolsModel.TypeFrequencyCheck = 0;
								toolsModel.ValueFrequency = 7;
							}
							else if (toolsModel.FrequencyCheck.Trim().ToUpper() == "NGÀY" || toolsModel.FrequencyCheck.Trim().ToUpper() == "DAILY" || toolsModel.FrequencyCheck.Trim().ToUpper() == "DAY")
							{
								toolsModel.TypeFrequencyCheck = 1;
								toolsModel.ValueFrequency = 1;
							}
							else if (toolsModel.FrequencyCheck.Trim().ToUpper() == "THÁNG" || toolsModel.FrequencyCheck.Trim().ToUpper() == "MONTH")
							{
								toolsModel.TypeFrequencyCheck = 2;
								toolsModel.ValueFrequency = 1;
							}
							else if (toolsModel.FrequencyCheck.Trim().ToUpper() == "NĂM" || toolsModel.FrequencyCheck.Trim().ToUpper() == "YEAR")
							{
								toolsModel.TypeFrequencyCheck = 3;
								toolsModel.ValueFrequency = 1;
							}
							toolsModel.CreateDate = TextUtils.ToDate(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));

							List<string> lstCommand = new List<string>();
							if (!chkWorking.Checked)
							{
								if (toolsModel.ID == 0)
								{
									toolsModel.ID = (int)ToolsBO.Instance.Insert(toolsModel);
									//Thêm các hạng mục check theo nhóm sản phẩm
									string sqlInitWorking = string.Format("EXEC dbo.CreateToolsWorking_ByGroupID {0},{1}", groupID, toolsModel.ID);
									lstCommand.Add(sqlInitWorking);
								}
								else
								{
									ToolsBO.Instance.Update(toolsModel);
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
										if (KnifeGroupCode.ToLower() != groupCode1.ToLower()) continue;
										string stepCode = arrW[1].Trim();
										string sortOrder = arrW[2].Trim();

										string sql = getSqlGroup(value, min, max, sortOrder, stepCode, toolsModel.ID, groupCode1);
										lstCommand.Add(sql);
									}

									if (arrW.Length == 2)//chỉ chứa mã công đoạn, stt mục kiểm tra
									{
										string stepCode = arrW[0].Trim();

										string sortOrder = arrW[1].Trim();
										string sql = getSql(value, min, max, sortOrder, stepCode, toolsModel.ID);
										lstCommand.Add(sql);
									}
								}
							}

							if (lstCommand.Count > 0)
							{
								string sqlRun = string.Join(";", lstCommand);
								TextUtilsHP.ExcuteSQL(sqlRun);
							}
						}
						else
						{
							KnifeDetailListModel knifeDetailList = new KnifeDetailListModel();
							//var lstObject = lstProduct.FirstOrDefault(o => ((ProductModel)o).ProductCode.Trim().ToLower() == productCode.Trim().ToLower());
							ArrayList arrP = KnifeDetailListBO.Instance.FindByAttribute("KnifeCode", KnifeCode);
							int OldGroupID = 0;
							if (arrP.Count > 0)
							{
								knifeDetailList = (KnifeDetailListModel)arrP[0];
								OldGroupID = knifeDetailList.KnifeGroupID;
								if (OldGroupID != groupID)
								{
									//Xóa trường cũ đi và thêm trường mới vào
									KnifeJigWorkingBO.Instance.DeleteByAttribute("KnifeDetailListID", knifeDetailList.ID);
									KnifeStepBO.Instance.DeleteByAttribute("KnifeDetailListID", knifeDetailList.ID);
									KnifeDetailListBO.Instance.Delete(knifeDetailList.ID);
									knifeDetailList.ID = 0;
								}
							}
							#region SetValue
							string valueString = "";
							knifeDetailList.KnifeGroupID = groupID;
							//knifeDetailList.ProductGroupCode = KnifeGroupCode;
							knifeDetailList.KnifeCode = KnifeCode;
							knifeDetailList.KnifeName = Lib.ToString(grvData.GetRowCellValue(i, "F3"));//Tên dụng cụ
							//knifeDetailList.WorkerID = Lib.ToString(grvData.GetRowCellValue(i, "F4"));//Người nhập
																									   //knifeDetailList.RatioCode = Lib.ToString(grvData.GetRowCellValue(i, "F5"));//Bộ phận
							knifeDetailList.Description = Lib.ToString(grvData.GetRowCellValue(i, "F6"));//Mô tả				
							knifeDetailList.STD = Lib.ToInt(grvData.GetRowCellValue(i, "F7"));//STD
							knifeDetailList.ATC = Lib.ToInt(grvData.GetRowCellValue(i, "F8"));//ACT
							knifeDetailList.DateOld = Lib.ToDate2(grvData.GetRowCellValue(i, "F11"));// khai báo ngày để hiện của lịch sử 
							knifeDetailList.FrequencyCheck = Lib.ToString(grvData.GetRowCellValue(i, "F12")); // tần suất kiểm tra theo ngày tháng năm
							 if(knifeDetailList.FrequencyCheck.Trim().ToUpper() == "TUẦN" || knifeDetailList.FrequencyCheck.Trim().ToUpper() == "WEEK")
							{
								knifeDetailList.TypeFrequencyCheck = 0;
								knifeDetailList.ValueFrequency = 7;
							}
							else if (knifeDetailList.FrequencyCheck.Trim().ToUpper() == "NGÀY" || knifeDetailList.FrequencyCheck.Trim().ToUpper() == "DAILY" || knifeDetailList.FrequencyCheck.Trim().ToUpper() == "DAY")
							{
								knifeDetailList.TypeFrequencyCheck = 1;
								knifeDetailList.ValueFrequency = 1;
							}
							else if(knifeDetailList.FrequencyCheck.Trim().ToUpper() == "THÁNG" || knifeDetailList.FrequencyCheck.Trim().ToUpper() == "MONTH")
							{
								knifeDetailList.TypeFrequencyCheck = 2;
								knifeDetailList.ValueFrequency = 1;
							}	
							else if (knifeDetailList.FrequencyCheck.Trim().ToUpper() == "NĂM" || knifeDetailList.FrequencyCheck.Trim().ToUpper() == "YEAR")
							{
								knifeDetailList.TypeFrequencyCheck = 3;
								knifeDetailList.ValueFrequency = 1;
							}	
							knifeDetailList.Type = cboType.SelectedIndex - 1;
							knifeDetailList.Status = true;
							knifeDetailList.ImportedDate = DateTime.Now;
							#endregion

							List<string> lstCommand = new List<string>();
							if (!chkWorking.Checked)
							{
								if (knifeDetailList.ID == 0)
								{
									knifeDetailList.ID = (int)KnifeDetailListBO.Instance.Insert(knifeDetailList);
									//Thêm các hạng mục check theo nhóm sản phẩm
									string sqlInitWorking = string.Format("EXEC dbo.CreateKnifeJigWorking_ByGroupID {0},{1}", groupID, knifeDetailList.ID);
									lstCommand.Add(sqlInitWorking);
								}
								else
								{
									KnifeDetailListBO.Instance.Update(knifeDetailList);
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
								string Frequency = "";
								if (value.Contains(spt))
								{
									string[] arrValue = value.Split(spt);
									min = arrValue[0].Trim();
									max = arrValue[1].Trim();
								}
								else if (value.Contains(dau)) //Check có dấu / thì là tần suất ví dụ "1/5"
								{
									Frequency = value;
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
										if (KnifeGroupCode.ToLower() != groupCode1.ToLower()) continue;

										string stepCode = arrW[1].Trim();
										string sortOrder = arrW[2].Trim();

										if (Frequency != "")
										{
											string sqlFrequency = getSqlGroupFrequency(Frequency, sortOrder, stepCode, knifeDetailList.ID, groupCode1);
											lstCommand.Add(sqlFrequency);
										}
										else
										{
											string sql = getSqlGroup(value, min, max, sortOrder, stepCode, knifeDetailList.ID, groupCode1);
											lstCommand.Add(sql);
										}

									}

									if (arrW.Length == 2)//chỉ chứa mã công đoạn, stt mục kiểm tra
									{
										string stepCode = arrW[0].Trim();

										string sortOrder = arrW[1].Trim();

										if (Frequency != "")
										{
											string sqlFrequency = getSqlFrequency(Frequency, sortOrder, stepCode, knifeDetailList.ID);
											lstCommand.Add(sqlFrequency);
										}
										else
										{
											string sql = getSql(value, min, max, sortOrder, stepCode, knifeDetailList.ID);
											lstCommand.Add(sql);
										}
									}
								}
							}

							if (lstCommand.Count > 0)
							{
								string sqlRun = string.Join(";", lstCommand);
								TextUtilsHP.ExcuteSQL(sqlRun);
							}
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
			});

		}
		private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show(start.ToString() + " - " + DateTime.Now.ToString());
			enableControl(true);
		}

		private void frmImportToolExcel_Load(object sender, EventArgs e)
		{

		}
	}
}
