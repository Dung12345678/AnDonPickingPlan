using HP.Utils;
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
	public partial class frmImportGoodsExcelH : Form
	{
		private DataSet ds;
		DataTable dt;
		public frmImportGoodsExcelH()
		{
			InitializeComponent();
		}
		private void frmImportProductExcelH_Load(object sender, EventArgs e)
		{

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
				progressBar1.Maximum = grvData.RowCount;
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

		private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			txtRate.Text = "";
			int rowCount = grvData.RowCount;
			for (int i = 0; i < rowCount; i++)
			{
				try
				{
					if (i < 1) continue;
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i, rowCount); }));
					string Code = TextUtilsHP.ToString(grvData.GetRowCellValue(i, "F1"));// Code
					string Name = TextUtilsHP.ToString(grvData.GetRowCellValue(i, "F2")); // Name
					if (string.IsNullOrEmpty(Code) || string.IsNullOrEmpty(Name))
					{
						continue;
					}
					Expression exp1 = new Expression("Code", Code);
					Expression exp2 = new Expression("Name", Name);
					ArrayList arr = GoodsBO.Instance.FindByExpression(exp1.And(exp2));
					GoodsModel goodsModel = new GoodsModel();
					if (arr.Count > 0)
					{
						// update
						goodsModel.Code = Code;
						goodsModel.Name = Name;
						GoodsBO.Instance.Update(goodsModel);
					}
					else
					{
						//Insert
						goodsModel.Code = Code;
						goodsModel.Name = Name;
						goodsModel.CreateDate = DateTime.Now;
						GoodsBO.Instance.Insert(goodsModel);
						
					}

					#region SetValue

					#endregion
				}
				catch
				{
					MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i);
				}
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


	}
}
