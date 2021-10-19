using DevExpress.XtraCharts;
using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmKnifeInformation : _Forms
	{
		public string _KnifeCode = "";
		public string _Goods = "";
		public frmKnifeInformation()
		{
			InitializeComponent();
			
		}
		private void frmDataAccess_Load(object sender, EventArgs e)
		{
			LoadKnife();
			LoadGridData();
			LoadDataToChart();
		}
		void LoadDataToChart()
		{
			DataSet ds = TextUtilsHP.GetListDataFromSP("spGetChartProductHistoryData", new string[] { "@knifeCode" }, new object[] { _KnifeCode });
			DataTable data = ds.Tables[0];
			DataTable dtRealTime = ds.Tables[1];

			DataTable table = new DataTable();
			DataRow row = null;
			table.Columns.Add("CDay", typeof(string));
			table.Columns.Add("Quantity", typeof(int));

			for (int i = 0; i < data.Rows.Count; i++)
			{
				row = table.NewRow();
				row["CDay"] = TextUtilsHP.ToString(data.Rows[i]["CDay"]);
				row["Quantity"] = TextUtilsHP.ToString(data.Rows[i]["Quantity"]);
				table.Rows.Add(row);
			}
			for (int i = 0; i < dtRealTime.Rows.Count; i++)
			{
				row = table.NewRow();
				row["CDay"] = TextUtilsHP.ToString(dtRealTime.Rows[0]["CDay"]);
				row["Quantity"] = TextUtilsHP.ToString(dtRealTime.Rows[0]["Quantity"]);
				table.Rows.Add(row);
			}

			chartControl1.Series[0].DataSource = table;
			chartControl1.Series[0].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[0].ArgumentDataMember = "CDay";
			chartControl1.Series[0].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "Quantity" });

		}
		/// <summary>
		/// Hiển thị lịch sử Dao khi gia công
		/// </summary>
		void LoadGridData()
		{
			DataTable dt = TextUtilsHP.Select($"SELECT * FROM  [HypoidPinion].[dbo].[KnifeProcessedList] WHERE KnifeCode = N'{_KnifeCode.Trim()}' ORDER BY ID DESC");
			DataTable dtt = dt.Select("Worker<>'PGN'").CopyToDataTable();
			grdData.DataSource = dtt;
			//Lấy dữ liệu ở PGN hiển thị lên 
			DataRow[] drPGN = dt.Select($"GoodsCode='{_Goods}'");
			if (drPGN.Count() <= 0) return;
			txtSTTGC.Text = TextUtilsHP.ToString(drPGN[0]["Quantity"]);
		}
		void LoadKnife()
		{
			try
			{
				DataTable dt = TextUtilsHP.Select($"SELECT Top 1 * FROM  [HypoidPinion].[dbo].[KnifeDetailList] WHERE KnifeCode = N'{_KnifeCode.Trim()}' AND Type = 0");
				if (dt.Rows.Count <= 0)
				{
					MessageBox.Show("Không tìm thấy mã Dao");
					return;
				}
				else
				{
					txtCode.Text = TextUtilsHP.ToString(dt.Rows[0]["KnifeCode"]);
					txtCurrentATC.Text = TextUtilsHP.ToString(dt.Rows[0]["CurrentATC"]);
					txtCurrentSTD.Text = TextUtilsHP.ToString(dt.Rows[0]["CurrentSTD"]);
					txtSTD.Text = TextUtilsHP.ToString(dt.Rows[0]["STD"]);
					txtRemainQty.Text = TextUtilsHP.ToString(dt.Rows[0]["RemainQty"]);
					txtTotalNumber.Text = TextUtilsHP.ToString(dt.Rows[0]["TotalNumber"]);
					txtDateTime.Text = TextUtilsHP.ToString(dt.Rows[0]["UpdatedDate"]);
				}
			}
			catch (Exception)
			{
			}
		}

		private void chartControl1_CustomDrawSeriesPoint(object sender, DevExpress.XtraCharts.CustomDrawSeriesPointEventArgs e)
		{
			if (string.Equals(e.SeriesPoint.Argument, "REAL TIME"))
			{
				e.SeriesDrawOptions.Color = Color.Pink;
				//e.SeriesDrawOptions.ActualColor2 = Color.Red;
			}
		}

		private void chartControl1_CustomDrawSeries(object sender, CustomDrawSeriesEventArgs e)
		{

		}
	}
}
