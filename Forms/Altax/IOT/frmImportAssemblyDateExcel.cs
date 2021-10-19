using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using ExcelDataReader;

namespace BMS
{
	public partial class frmImportAssemblyDateExcelAltax : _Forms
	{
		private DataSet ds;
		public frmImportAssemblyDateExcelAltax()
		{
			InitializeComponent();
		}
		DateTime start;
		private void frmImportAssemblyDateExcel_Load(object sender, EventArgs e)
		{
			btnSave.Enabled = false;
		}
		void enableControl(bool enable)
		{
			btnSave.Enabled = enable;
			grdData.Enabled = enable;
			cboSheet.Enabled = enable;
			btnBrowse.Enabled = enable;
		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			if (backgroundWorker1.IsBusy)
			{
				backgroundWorker1.CancelAsync();
			}
			else
			{
				progressBar1.Minimum = 1;
				progressBar1.Maximum = grvData.RowCount - 1;
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

					dtt.Columns.Add("DATEF2", typeof(DateTime));
					for (int i = 0; i < dtt.Rows.Count; i++)
					{
						if (i < 1
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F2"]).Trim() != "")
							dtt.Rows[i]["DATEF2"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F2"]));
					}

					dtt.Columns.Add("DATEF3", typeof(DateTime));
					for (int i = 0; i < dtt.Rows.Count; i++)
					{
						if (i < 1
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F3"]).Trim() != "")
							dtt.Rows[i]["DATEF3"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F3"]));
					}
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
					dt.Columns.Add("DATEF2", typeof(DateTime));
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						if (i < 1
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F1"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F2"]).Trim() != "")
							dt.Rows[i]["DATEF2"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F2"]));
					}

					dt.Columns.Add("DATEF3", typeof(DateTime));
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						if (i < 1
							) continue;
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F1"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F3"]).Trim() != "")
							dt.Rows[i]["DATEF3"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F3"]));
					}



					grdData.DataSource = dt;
					grvData.PopulateColumns();
					grvData.BestFitColumns();
					grdData.Focus();

					//grvData.Columns[3].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
					//grvData.Columns[3].DisplayFormat.FormatString = "dd/MM/yyyy";
				}
				catch (Exception ex)
				{
					TextUtils.ShowError(ex);
					grdData.DataSource = null;
				}
			}
		}

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			txtRate.Text = "";

			int rowCount = grvData.RowCount;
			for (int i = 0; i < rowCount; i++)
			{
				try
				{
					if (i < 1) continue;
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i, rowCount - 1); }));

					string ShipTo = Lib.ToString(grvData.GetRowCellValue(i, "F1")); // Đích	
					DateTime DeliveryDate;
					DateTime AssemblyDate;
					//Kiểm tra nếu đích trống thì next
					if (string.IsNullOrEmpty(ShipTo))
					{
						continue;
					}

					DeliveryDate = Lib.ToDate3(grvData.GetRowCellValue(i, "F2")); // Ngày xuất hàng
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F2")) == new DateTime(1950, 1, 1))
					{
						DeliveryDate = Lib.ToDate3(grvData.GetRowCellValue(i, "DATEF2"));
					}

					AssemblyDate = Lib.ToDate3(grvData.GetRowCellValue(i, "F3"));//Ngày lắp ráp	
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F3")) == new DateTime(1950, 1, 1))
					{
						AssemblyDate = Lib.ToDate3(grvData.GetRowCellValue(i, "DATEF3"));
					}

					TextUtils.ExcuteProcedure("spUpdateAssemblyDate", new string[] { "@ShipTo", "@DeliveryDate", "@AssemblyDate" }, new object[] { ShipTo, DeliveryDate, AssemblyDate });
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
	}
}

