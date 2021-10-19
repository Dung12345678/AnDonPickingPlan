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
using System.Data.OleDb;
using System.Globalization;
using ST.Business;
using ST.Model;
using ST.Utils;

namespace BMS
{

	public partial class frmImPortOrderPart : _Forms
	{
		private DataSet ds;
		private Thread _threadUpdateDate;
		public frmImPortOrderPart()
		{
			InitializeComponent();
		}
		private void frmImportProductionPlanExcel_Load(object sender, EventArgs e)
		{
			btnSave.Enabled = true;
			//thread save UpdateDate
			//UpdateDate();
		}
		void enableControl(bool enable)
		{
			btnSave.Enabled = enable;
			grdData.Enabled = enable;
			cboSheet.Enabled = enable;
			btnBrowse.Enabled = enable;
		}

		DateTime start;
		void UpdateDate()
		{
			try
			{
				//grvData.Columns.Clear();
				string path = "U:\\Production Control\\PGN EXPORT\\VN_EXPORT_ACCESSORY_TO_ASSEMBLING.CSV";

				DataTable dtt = GetDataTableFromCsv(path, checkisFirstRowHeader.Checked);
				for (int i = 0; i < dtt.Columns.Count; i++)
				{
					dtt.Columns[i].ColumnName = "F" + (i + 1);
				}

				if (dtt.Rows.Count <= 0) return;
				dtt.Columns.Add("DATEF9", typeof(DateTime));
				for (int i = 1; i < dtt.Rows.Count; i++)
				{
					if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F2"])))
					{
						continue;
					}
					if (Lib.ToString(dtt.Rows[i]["F9"]).Trim() != "")
						dtt.Rows[i]["DATEF9"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F9"]));
				}
				grdData.DataSource = dtt;
				Save();

			}
			catch
			{

			}
		}
		async void Save()
		{
			Task task = Task.Factory.StartNew(() =>
			{
				int rowCount = grvData.RowCount;
				for (int i = 1; i < rowCount; i++)
				{
					try
					{
						string _ordercode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
						string Shelf = Lib.ToString(grvData.GetRowCellValue(i, "F6"));

						string _ArticleID = Lib.ToString(grvData.GetRowCellValue(i, "F3"));
						//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
						if (string.IsNullOrEmpty(_ordercode) || string.IsNullOrEmpty(_ArticleID))
						{
							continue;
						}
						Expression exp2 = new Expression("AssemblyOrder", _ordercode);
						Expression exp1 = new Expression("ArticleID", _ArticleID);
						ArrayList arr = OrderPartBO.Instance.FindByExpression(exp1.And(exp2));
						OrderPartModel orderPart = new OrderPartModel();

						#region SetValue
						string a = "";
						orderPart.OrderCode = _ordercode;
						orderPart.ArticleID = _ArticleID;
						orderPart.Cnt = Lib.ToInt(grvData.GetRowCellValue(i, "F2"));
						orderPart.Description = Lib.ToString(grvData.GetRowCellValue(i, "F4"));
						orderPart.Qty = Lib.ToInt(Lib.ToDouble(grvData.GetRowCellValue(i, "F5")));
						orderPart.Shelf = Shelf;
						orderPart.Location = Lib.ToString(grvData.GetRowCellValue(i, "F7"));
						orderPart.Lot = Lib.ToString(grvData.GetRowCellValue(i, "F8"));
						orderPart.CreateAt = Lib.ToDate3(grvData.GetRowCellValue(i, "F9"));//date
						if (Lib.ToDate3(grvData.GetRowCellValue(i, "F9")) == new DateTime(1950, 1, 1))
						{
							orderPart.CreateAt = Lib.ToDate2(grvData.GetRowCellValue(i, "DATEF9").ToString());
						}
						orderPart.Userr = Lib.ToString(grvData.GetRowCellValue(i, "F10")); // đích
						orderPart.OrderCodeAndCnt = _ordercode + orderPart.Cnt;
						#endregion

						if (arr.Count > 0)
						{
							//orderPart.UpdatedAt = DateTime.Now;
							for (int j = 0; j < arr.Count; j++)
							{
								orderPart.UpdateDate = DateTime.Now;
								orderPart.ID = (arr[j] as OrderPartModel).ID;
								OrderPartBO.Instance.Update(orderPart);
							}
						}

						else
						{
							orderPart.CreateDate = DateTime.Now;
							orderPart.ID = (int)OrderPartBO.Instance.Insert(orderPart);
						}

					}
					catch (Exception ex)
					{
						MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i + Environment.NewLine + ex.ToString());
					}
				}
			});
			await task;
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
			DataTable dtt = GetDataTableFromCsv(btnBrowse.Text.Trim(), checkisFirstRowHeader.Checked);
			for (int i = 0; i < dtt.Columns.Count; i++)
			{
				dtt.Columns[i].ColumnName = "F" + (i + 1);
			}

			if (dtt.Rows.Count <= 0) return;
			dtt.Columns.Add("DATEF9", typeof(DateTime));
			for (int i = 1; i < dtt.Rows.Count; i++)
			{
				if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F2"])))
				{
					continue;
				}
				if (Lib.ToString(dtt.Rows[i]["F9"]).Trim() != "")
					dtt.Rows[i]["DATEF9"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F9"]));
			}

			grdData.DataSource = dtt;
			grvData.PopulateColumns();
			grvData.BestFitColumns();
			grdData.Focus();
			//return;
		}
		public static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
		{
			string header = isFirstRowHeader ? "Yes" : "No";

			string pathOnly = Path.GetDirectoryName(path);
			string fileName = Path.GetFileName(path);

			string sql = @"SELECT * FROM [" + fileName + "]";

			using (OleDbConnection connection = new OleDbConnection(
					  @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
					  ";Extended Properties=\"Text;HDR=" + header + "\""))
			using (OleDbCommand command = new OleDbCommand(sql, connection))
			using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
			{
				DataTable dataTable = new DataTable();
				dataTable.Locale = CultureInfo.CurrentCulture;
				adapter.Fill(dataTable);
				return dataTable;
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

					dtt.Columns.Add("DATEF9", typeof(DateTime));
					for (int i = 1; i < dtt.Rows.Count; i++)
					{
						if (string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dtt.Rows[i]["F2"])))
						{
							continue;
						}
						if (Lib.ToString(dtt.Rows[i]["F9"]).Trim() != "")
							dtt.Rows[i]["DATEF9"] = DateTime.FromOADate(TextUtils.ToDouble(dtt.Rows[i]["F9"]));
					}

					//var tablename = cboSheet.SelectedItem.ToString();

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
					dt.Columns.Add("DATEF9", typeof(DateTime));
					for (int i = 1; i < dt.Rows.Count; i++)
					{
						if (string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F1"])) || string.IsNullOrEmpty(Lib.ToString(dt.Rows[i]["F2"])))
						{
							continue;
						}
						if (Lib.ToString(dt.Rows[i]["F9"]).Trim() != "")
							dt.Rows[i]["DATEF9"] = DateTime.FromOADate(TextUtils.ToDouble(dt.Rows[i]["F9"]));
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

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			txtRate.Text = "";

			int rowCount = grvData.RowCount;
			for (int i = 1; i < rowCount; i++)
			{
				try
				{
					progressBar1.Invoke((Action)(() => { progressBar1.Value = i; }));
					txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i, rowCount); }));
					string _ordercode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
					string Shelf = Lib.ToString(grvData.GetRowCellValue(i, "F6"));

					string _ArticleID = Lib.ToString(grvData.GetRowCellValue(i, "F3"));
					//Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
					if (string.IsNullOrEmpty(_ordercode) || string.IsNullOrEmpty(Shelf))
					{
						continue;
					}
					Expression exp2 = new Expression("OrderCode", _ordercode);
					Expression exp1 = new Expression("ArticleID", _ArticleID);
					ArrayList arr = OrderPartBO.Instance.FindByExpression(exp1.And(exp2));
					OrderPartModel orderPart = new OrderPartModel();

					#region SetValue
					string a = "";
					orderPart.OrderCode = _ordercode;
					orderPart.ArticleID = _ArticleID;
					orderPart.Cnt = Lib.ToInt(grvData.GetRowCellValue(i, "F2"));
					orderPart.Description = Lib.ToString(grvData.GetRowCellValue(i, "F4"));

					orderPart.Qty = Lib.ToInt(Lib.ToDouble(grvData.GetRowCellValue(i, "F5")));
					orderPart.Shelf = Shelf;
					orderPart.Location = Lib.ToString(grvData.GetRowCellValue(i, "F7"));
					orderPart.Lot = Lib.ToString(grvData.GetRowCellValue(i, "F8"));
					orderPart.CreateAt = Lib.ToDate3(grvData.GetRowCellValue(i, "F9"));//date
					if (Lib.ToDate3(grvData.GetRowCellValue(i, "F9")) == new DateTime(1950, 1, 1))
					{
						orderPart.CreateAt = Lib.ToDate2(grvData.GetRowCellValue(i, "DATEF9").ToString());
					}
					orderPart.Userr = Lib.ToString(grvData.GetRowCellValue(i, "F10")); // đích
					orderPart.OrderCodeAndCnt = _ordercode + orderPart.Cnt;
					#endregion

					if (arr.Count > 0)
					{
						//orderPart.UpdatedAt = DateTime.Now;
						
						for (int j = 0; j < arr.Count; j++)
						{
							orderPart.CreateDate = (arr[j] as OrderPartModel).CreateDate;
							orderPart.UpdateDate = DateTime.Now;
							orderPart.ID = (arr[j] as OrderPartModel).ID;
							
							OrderPartBO.Instance.Update(orderPart);
						}
					}

					else
					{
						orderPart.CreateDate = DateTime.Now;
						orderPart.ID = (int)OrderPartBO.Instance.Insert(orderPart);
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
			//enableControl(true);
		}

		private void btnBrowse_BindingContextChanged(object sender, EventArgs e)
		{

		}
	}
}
