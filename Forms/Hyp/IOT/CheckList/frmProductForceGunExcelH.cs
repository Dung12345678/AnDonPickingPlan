using IE.Business;
using IE.Model;
using IE.Utils;
using DevExpress.XtraEditors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using System.Diagnostics;

namespace BMS
{
	public partial class frmProductForceGunExcelH : _Forms
	{
		private DataSet ds;
		public frmProductForceGunExcelH()
		{
			InitializeComponent();
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

		private void frmImportCheckForceExcel_Load(object sender, EventArgs e)
		{
            //ArrayList
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

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
                     
            int rowCount = grvData.RowCount;
            int colCount = grvData.Columns.Count;
			ProductCheckForceModel model = new ProductCheckForceModel();
            for (int i = 0; i < rowCount; i++)
            {
				try
				{
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i + 1; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i + 1, rowCount); }));

					// update CheckForce
					if (string.IsNullOrEmpty(grvData.GetRowCellValue(i, "F1").ToString().Trim()) || string.IsNullOrEmpty(grvData.GetRowCellValue(i, "F11").ToString().Trim()))
						continue;
					Expression expression = new Expression("ProductCode", Lib.ToString(grvData.GetRowCellValue(i, "F1")));
					ArrayList lst = ProductBO.Instance.FindByExpression(expression);
					if (lst.Count <= 0)
						continue;
					string sql = string.Format("Select ID from ProductCheckForce Where ProductCode = '{0}' And ProductStepCode = '{1}'",
						Lib.ToString(grvData.GetRowCellValue(i, "F1")), Lib.ToString(grvData.GetRowCellValue(i, "F13")));
					SqlDataReader reader = SqlHelper.ExecuteReader(DBUtils.GetDBConnectionString(), CommandType.Text, sql);
					if (reader.Read())
					{
						model = ProductCheckForceBO.Instance.FindByPK(long.Parse(reader[0].ToString())) as ProductCheckForceModel;
						string[] CheckForce = Lib.ToString(grvData.GetRowCellValue(i, "F5")).Split('~');
						model.CheckForceMin = TextUtils.ToDecimal(CheckForce[0]);
						model.CheckForceMax = TextUtils.ToDecimal(CheckForce[1]);
						model.GunNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F6"));
						model.JobNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F7"));
						model.QtyBuLong = Lib.ToInt(grvData.GetRowCellValue(i, "F8"));	
						model.ShootNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F9"));
						model.QtyOcBanGa = Lib.ToInt(grvData.GetRowCellValue(i, "F10"));
						model.QtyOcBanThat = Lib.ToInt(grvData.GetRowCellValue(i, "F11"));
						model.SetForce = Lib.ToDecimal(grvData.GetRowCellValue(i, "F12"));
						ProductCheckForceBO.Instance.Update(model);
					}
					else
					{
						model.ProductCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
						string[] CheckForce = Lib.ToString(grvData.GetRowCellValue(i, "F5")).Split('~');
						model.CheckForceMin = TextUtils.ToDecimal(CheckForce[0]);
						model.CheckForceMax = TextUtils.ToDecimal(CheckForce[1]);
						model.GunNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F6"));
						model.JobNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F7"));
						model.QtyBuLong = Lib.ToInt(grvData.GetRowCellValue(i, "F8"));
						model.ShootNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F9"));
						model.QtyOcBanGa = Lib.ToInt(grvData.GetRowCellValue(i, "F10"));
						model.QtyOcBanThat = Lib.ToInt(grvData.GetRowCellValue(i, "F11"));
						model.SetForce = Lib.ToDecimal(grvData.GetRowCellValue(i, "F12"));
						model.ProductStepCode = Lib.ToString(grvData.GetRowCellValue(i, "F13"));
						ProductCheckForceBO.Instance.Insert(model);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(i.ToString() + Environment.NewLine + ex.ToString());
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
