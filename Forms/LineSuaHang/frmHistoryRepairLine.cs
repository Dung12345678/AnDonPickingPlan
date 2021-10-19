using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System.Reflection;
using System.Threading;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using System.Diagnostics;
using DevExpress.XtraGrid.Localization;
using iTextSharp;
using iTextSharp.text.pdf;
using DevExpress.Utils;
using IOEx;
using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.XtraGrid.Columns;
using System.Web;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace BMS
{
	public partial class frmHistoryRepairLine : _Forms
	{
		private string _pathFileConfigUpdate = Path.Combine(Application.StartupPath, "Download.txt");
		public frmHistoryRepairLine()
		{
			InitializeComponent();
		}

		private void frmHistoryCheck_Load(object sender, EventArgs e)
		{
			dtpFromDate.Value = DateTime.Now;
			dtpEndDate.Value = DateTime.Now;
		}
		List<string> ErrColumnsName = new List<string>();
		void loadData()
		{
			ErrColumnsName = new List<string>();
			DateTime from = new DateTime(dtpFromDate.Value.Year, dtpFromDate.Value.Month, dtpFromDate.Value.Day, 0, 0, 0);
			DateTime to = new DateTime(dtpEndDate.Value.Year, dtpEndDate.Value.Month, dtpEndDate.Value.Day, 23, 59, 59);

			DataTable dtFind = TextUtilsStock.LoadDataFromSP("spGetLineRepairData", "A"
				, new string[] { "@FromDate", "@EndDate", "@TextFind" }
				, new object[] { from, to, txtFindText.Text.Trim() });
			if (grvData.Columns.Count < dtFind.Columns.Count)
			{
				//Thêm cột theo dt
				for (int i = 0; i < dtFind.Columns.Count - 6; i++)
				{
					GridColumn unboundColumn = grvData.Columns.Add();
					unboundColumn.ShowUnboundExpressionMenu = true;
					unboundColumn.FieldName = dtFind.Columns[dtFind.Columns.Count - 7 + i].ColumnName;
					unboundColumn.Name = "col" + dtFind.Columns[dtFind.Columns.Count - 7 + i].ColumnName;
					unboundColumn.Caption = dtFind.Columns[dtFind.Columns.Count - 7 + i].ColumnName;
					unboundColumn.AppearanceHeader.Options.UseTextOptions = true;
					unboundColumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
					unboundColumn.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
					unboundColumn.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
					unboundColumn.AppearanceCell.Options.UseTextOptions = true;
					unboundColumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
					unboundColumn.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
					unboundColumn.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
					unboundColumn.AppearanceCell.ForeColor = Color.Blue;
					unboundColumn.AppearanceCell.Font = new Font(unboundColumn.AppearanceCell.Font, FontStyle.Bold | FontStyle.Underline);
					unboundColumn.AppearanceHeader.Font = new Font("Tahoma", 8, FontStyle.Bold);
					unboundColumn.ColumnEdit = repositoryItemMemoEdit1;
					unboundColumn.Visible = true;
					unboundColumn.Width = 150;
					unboundColumn.VisibleIndex = grvData.Columns.Count + i;
					unboundColumn.OptionsColumn.AllowFocus = false;
					DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit txt = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
					txt.AllowHtmlDraw = DefaultBoolean.True;
					txt.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
					unboundColumn.ColumnEdit = txt;

				}
			}
			//Hiển thị view click vào sẽ view ảnh lên
			//string filename = Path.Combine(_pathImage, order + "_" + stt + "_" + qtyLap + "_" + vtriAnh + ".jpg");
			//if (File.Exists(filename))
			//{
			//	grvData.Rows[i].Cells["colImage" + h].Value = "View";
			//}
			grdData.DataSource = dtFind;
			List<string> defaultColumns = new List<string> { "WorkerCode", "CreateDate", "PartCode", "QrCode", "OrderCode", "NamePart" };
			if (dtFind != null)
			{
				foreach (DataColumn col in dtFind.Columns)
					if (!defaultColumns.Contains(col.ColumnName))
					{

						ErrColumnsName.Add(col.ColumnName);
					}
			}

		}
		private void btnFind_Click(object sender, EventArgs e)
		{
			loadData();
		}
		private void txtFindText_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				loadData();
			}
		}
		private void btnExportExecl_Click(object sender, EventArgs e)
		{
			try
			{
				if (grvData.RowCount > 0)
				{
					FolderBrowserDialog od = new FolderBrowserDialog();
					if (od.ShowDialog() == DialogResult.OK)
					{
						TextUtils.ExportExcel(grvData, od.SelectedPath, string.Format("Line_Sửa_Hàng_{0}", DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private void grvData_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
		{
			if (ErrColumnsName != null && ErrColumnsName.Contains(e.Column.FieldName))
			{
				e.DisplayText = "View";
			}
		}

		private void grvData_MouseMove(object sender, MouseEventArgs e)
		{
			GridHitInfo hi = grvData.CalcHitInfo(e.Location);
			if (hi != null && hi.Column != null && ErrColumnsName != null && ErrColumnsName.Contains(hi.Column.FieldName))
			{
				grdData.Cursor = Cursors.Hand;
			}
			else
			{
				grdData.Cursor = Cursors.Default;
			}
		}

		private void grvData_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
		{
			if (TextUtils.ToString(e.CellValue) != "")
				Process.Start($"{TextUtils.ToString(e.CellValue)}");
		}
	}
}
