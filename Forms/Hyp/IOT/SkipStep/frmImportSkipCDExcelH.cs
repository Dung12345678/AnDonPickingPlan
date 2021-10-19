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
using IE.Business;
using IE.Model;
using IE.Utils;
using DevExpress.Data.Filtering.Helpers;
using ExcelDataReader;
using System.Diagnostics;
using System.IO;

namespace BMS
{

	public partial class frmImportSkipCDExcelH : _Forms
	{
		private DataSet ds;
		public frmImportSkipCDExcelH()
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
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i - 1, rowCount - 1); }));
					string _productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
					//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
					if (string.IsNullOrEmpty(_productCode))
					{
						continue;
					}
					SkipCDModel skipCDModel = new SkipCDModel();
					Expression exp1 = new Expression("Product", _productCode);
					ArrayList arr = SkipCDBO.Instance.FindByExpression(exp1);
					if (arr.Count > 0)
					{
						skipCDModel = (SkipCDModel)arr[0];
					}

					#region SetValue
					skipCDModel.Product = _productCode;
					skipCDModel.CDSkip = TextUtils.ToString(grvData.GetRowCellValue(i, "F2"));
					#endregion

					if (arr.Count > 0)
					{
						SkipCDBO.Instance.Update(skipCDModel);
					}

					else
					{
						skipCDModel.ID = (int)SkipCDBO.Instance.Insert(skipCDModel);
						//lstProductionPlanBO.Add(product);
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
