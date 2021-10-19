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
using HP.Model;
using HP.Business;
using HP.Utils;

namespace BMS
{
	public partial class frmImportGoodsRegisterExcel : _Forms
	{
		private DataSet ds;
		public frmImportGoodsRegisterExcel()
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
			this.Invoke((MethodInvoker)delegate
			{


				DataSet ds = TextUtilsHP.GetListDataFromSP("spGetGoodsIDAndToolsID", new string[] { "@type" }, new object[] { cboType.SelectedIndex });
				DataTable dtGoods = ds.Tables[0];
				DataTable dtKnife = ds.Tables[1];

				for (int i = 0; i < rowCount; i++)
				{
					try
					{
						if (i < 1) continue;
						progressBar1.Invoke((Action)(() => { progressBar1.Value = i; }));
						txtRate.Invoke((Action)(() => { txtRate.Text = string.Format("{0}/{1}", i, rowCount - 1); }));

						string Goods = Lib.ToString(grvData.GetRowCellValue(i, "F1")); // Mã hàng
						if (Goods.Trim() == "") continue;
						string Toods = Lib.ToString(grvData.GetRowCellValue(i, "F2")); // Mã dụng cụ
						if (Toods.Trim() == "") continue;
						DataTable dt = dtGoods.Select($"Code='{Goods}'").CopyToDataTable();
						if (dt.Rows.Count <= 0) continue;
						int GoodsID = TextUtils.ToInt(dt.Rows[0]["ID"]);
						if (cboType.SelectedIndex == 3)
						{
							GoodsRegisterModel goodsRegisterModel = new GoodsRegisterModel();
							DataTable Tools1 = dtKnife.Select($"Code='{Toods}'").CopyToDataTable();
							if (Tools1.Rows.Count <= 0) continue;
							goodsRegisterModel.GoodsID = GoodsID;
							goodsRegisterModel.ToolsID = TextUtils.ToInt(Tools1.Rows[0]["ID"]);
							Expression exp1 = new Expression("GoodsID", GoodsID);
							Expression exp2 = new Expression("ToolsID", goodsRegisterModel.ToolsID);
							ArrayList array = RegisterBO.Instance.FindByExpression(exp1.And(exp2));
							if (array.Count > 0)
							{
								//goodsRegisterModel = (GoodsRegisterModel)array[0];
								//GoodsRegisterBO.Instance.Update(goodsRegisterModel);
							}
							else
							{
								GoodsRegisterBO.Instance.Insert(goodsRegisterModel);
							}
						}
						else
						{
							RegisterModel registerModel = new RegisterModel();
							DataTable Knife = dtKnife.Select($"KnifeCode='{Toods}'").CopyToDataTable();
							if (Knife.Rows.Count <= 0) continue;
							registerModel.ToolsID = TextUtils.ToInt(Knife.Rows[0]["ID"]);
							registerModel.GoodsID = GoodsID;
							Expression exp1 = new Expression("GoodsID", GoodsID);
							Expression exp2 = new Expression("ToolsID", registerModel.ToolsID);
							ArrayList array = RegisterBO.Instance.FindByExpression(exp1.And(exp2));
							if (array.Count > 0)
							{
								//registerModel = (RegisterModel)array[0];
								//RegisterBO.Instance.Update(registerModel);
							}
							else
							{
								RegisterBO.Instance.Insert(registerModel);
							}
						}
						if (string.IsNullOrEmpty(Goods) || string.IsNullOrEmpty(Toods))
						{
							continue;
						}

						int Check = Lib.ToInt(grvData.GetRowCellValue(i, "F3")); // Check 1 là check , 0 bỏ check

					}
					catch (Exception ex)
					{
						MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i + Environment.NewLine + ex.ToString());
					}

				}
			});
		}

		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show(start.ToString() + " - " + DateTime.Now.ToString());
			enableControl(true);
		}
	}
}

