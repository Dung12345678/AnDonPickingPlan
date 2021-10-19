using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Data.Filtering.Helpers;
using ExcelDataReader;
using System.Diagnostics;
using System.IO;
using BMS.Model;
using BMS.Business;
using BMS.Utils;

namespace BMS
{
	
	public partial class frmImportProductPlanExcel : _Forms
	{
		private DataSet ds;
		public frmImportProductPlanExcel()
		{
			InitializeComponent();
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
					btnSave.Enabled = true;
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
					btnSave.Enabled = true;
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

					dtt.Columns.Add("DATEF4", typeof(DateTime));
					for (int i = 0; i < dtt.Rows.Count; i++)
					{
						if (i < 5
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F5"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F7"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F4"]).Trim() != "")
							dtt.Rows[i]["DATEF4"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F4"]));
					}

					dtt.Columns.Add("DATEF18", typeof(DateTime));
					for (int i = 0; i < dtt.Rows.Count; i++)
					{
						if (i < 5
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F5"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F7"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F18"]).Trim() != "")
							dtt.Rows[i]["DATEF18"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F18"]));
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
					dt.Columns.Add("DATEF4", typeof(DateTime));
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						if (i < 5
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F5"])) || string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F7"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F4"]).Trim() != "")
							dt.Rows[i]["DATEF4"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F4"]));
					}

					dt.Columns.Add("DATEF18", typeof(DateTime));
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						if (i < 5
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F5"])) || string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F7"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F18"]).Trim() != "")
							dt.Rows[i]["DATEF18"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F18"]));
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

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			txtRate.Text = "";

			int rowCount = grvData.RowCount;
			for (int i = 0; i < rowCount; i++)
			{
				try
				{
					if (i < 5) continue;
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i - 2; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i - 2, rowCount - 3); }));
					string _productCode = Lib.ToString(grvData.GetRowCellValue(i, "F7"));
					string _ordercode = Lib.ToString(grvData.GetRowCellValue(i, "F5"));
					int _cnt= Lib.ToInt(grvData.GetRowCellValue(i, "F6"));
					//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
					if (string.IsNullOrEmpty(_ordercode) || string.IsNullOrEmpty(_productCode))
					{
						continue;
					}
					Expression exp2 = new Expression("OrderCode", _ordercode);
					Expression exp1 = new Expression("Cnt", _cnt);
					ArrayList arr = ProductionPlanBO.Instance.FindByExpression(exp1.And(exp2));
					ProductionPlanModel productionPlanModel = new ProductionPlanModel();

					#region SetValue
					string a = "";
					productionPlanModel.OrderCode = _ordercode;// Mô tả AssyOrder F5
					productionPlanModel.ProductCode = _productCode;// Mô tả ID F7
					productionPlanModel.AssemblyProduct = Lib.ToString(grvData.GetRowCellValue(i, "F1"));//Assembly order & CNT
					productionPlanModel.Stt = Lib.ToInt(grvData.GetRowCellValue(i, "F3"));//Mô tả số thứ tự
					productionPlanModel.JgDate = Lib.ToDate3(grvData.GetRowCellValue(i, "F4"));//date
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F4")) == new DateTime(1950, 1, 1))
					{
						productionPlanModel.JgDate = Lib.ToDate2(grvData.GetRowCellValue(i, "DATEF4").ToString());
					}

					productionPlanModel.Cnt =_cnt;//Mô tả CNT
					productionPlanModel.Description = Lib.ToString(grvData.GetRowCellValue(i, "F8"));//Description
					productionPlanModel.Qty = Lib.ToInt(grvData.GetRowCellValue(i, "F9"));
					productionPlanModel.SalesOrder = Lib.ToString(grvData.GetRowCellValue(i, "F10")); // Order
					productionPlanModel.ShipTo = Lib.ToString(grvData.GetRowCellValue(i, "F11")); // đích
					productionPlanModel.SVia = Lib.ToString(grvData.GetRowCellValue(i, "F12"));

					productionPlanModel.Code = Lib.ToString(grvData.GetRowCellValue(i, "F13"));
					productionPlanModel.OrderNew = Lib.ToString(grvData.GetRowCellValue(i, "F14"));
					productionPlanModel.ProductCodeNew = Lib.ToString(grvData.GetRowCellValue(i, "F16"));
					productionPlanModel.Note = Lib.ToString(grvData.GetRowCellValue(i, "F17"));
					productionPlanModel.RequestDate = Lib.ToDate3(grvData.GetRowCellValue(i, "F18")); // ngày khách hàng yêu cầu    
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F18")) == new DateTime(1950, 1, 1))
					{
						productionPlanModel.RequestDate = Lib.ToDate2(grvData.GetRowCellValue(i, "DATEF18"));
					}
					productionPlanModel.MaMoto = Lib.ToString(grvData.GetRowCellValue(i, "F21"));
					productionPlanModel.MaMoto1 = Lib.ToString(grvData.GetRowCellValue(i, "F22"));
					productionPlanModel.OrderCodeFull = productionPlanModel.OrderCode + productionPlanModel.Cnt;
					productionPlanModel.AssemblyDate = new DateTime(9998, 01, 01);
					#endregion

					if (arr.Count > 0)
					{
						productionPlanModel.UpdatedAt = DateTime.Now;
						for (int j = 0; j < arr.Count; j++)
						{
							productionPlanModel.ID = (arr[j] as ProductionPlanModel).ID;
							ProductionPlanBO.Instance.Update(productionPlanModel);
						}
					}

					else
					{
						productionPlanModel.UpdatedAt = DateTime.Now;
						productionPlanModel.CreatedAt = DateTime.Now;
						productionPlanModel.ID = (int)ProductionPlanBO.Instance.Insert(productionPlanModel);
					}

				}
				catch (Exception ex)
				{
					MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i + Environment.NewLine + ex.ToString());
				}
			}
		}

		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show(start.ToString() + " - " + DateTime.Now.ToString());
			enableControl(true);
		}

		private void frmImportProductionPlanExcel_Load(object sender, EventArgs e)
		{
			btnSave.Enabled = false;
		}
	}
}
