using DevExpress.XtraCharts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMS.Business;
using BMS.Model;
using System.Threading;

namespace BMS
{
	public partial class frmKnifeProcessedChart : _Forms
	{
		#region Variables
		private string _knifeCode;
		private int _refreshFlag = 0;
		private string _columnName = "";
		private int _exitFlag = 1;
		#endregion

		#region Methods
		public frmKnifeProcessedChart()
		{
			InitializeComponent();
		}


		void LoadDataToChartProduct()
		{

			DataSet dt = TextUtilsHP.GetListDataFromSP("spGetChartProductData", new string[] { "@numberOfProduct" }, new object[] { 0 });
			if (dt.Tables.Count < 2) return;
			//  Chong' timer khi refresh se chay tren 1 luong rieng biet nen se gay loi Index of out bound

			chartControl2.Series[0].DataSource = dt.Tables[0];
			chartControl2.Series[0].ArgumentScaleType = ScaleType.Auto;
			chartControl2.Series[0].ArgumentDataMember = "KnifeCode";
			chartControl2.Series[0].ValueScaleType = ScaleType.Numerical;
			chartControl2.Series[0].ValueDataMembers.AddRange(new string[] { "Quantity" });


			chartControl2.Series[1].DataSource = dt.Tables[1];
			chartControl2.Series[1].ArgumentScaleType = ScaleType.Auto;
			chartControl2.Series[1].ArgumentDataMember = "KnifeCode";
			chartControl2.Series[1].ValueScaleType = ScaleType.Numerical;
			chartControl2.Series[1].ValueDataMembers.AddRange(new string[] { "STD" });

			chartControl2.Series[2].DataSource = dt.Tables[1];
			chartControl2.Series[2].ArgumentScaleType = ScaleType.Auto;
			chartControl2.Series[2].ArgumentDataMember = "KnifeCode";
			chartControl2.Series[2].ValueScaleType = ScaleType.Numerical;
			chartControl2.Series[2].ValueDataMembers.AddRange(new string[] { "ATC" });
			GetDetailToChart(dt.Tables[0]);

		}

		void RefreshDummySeries()
		{
			DataTable dt = new DataTable();
			chartControl2.Series[3].DataSource = dt;
		}

		private async void GetDetailToChart(DataTable dt)
		{
			try
			{
				Task task = Task.Factory.StartNew(() =>
				{
					for (int i = 0; i < dt.Rows.Count; i++)
					{
						Thread.Sleep(1000);
						if (_exitFlag == 2) return;
						_exitFlag = 1;

						_columnName = TextUtilsHP.ToString(dt.Rows[i]["KnifeCode"]);
						chartControl1.Invoke((MethodInvoker)delegate
						{
							LoadDataToChartTime(_columnName);
							RefreshDummySeries();
						});
						_exitFlag = 0;
					}
				}
			);
				await task;
				_refreshFlag = 0;
			}
			catch
			{
				return;
			}

		}

		void LoadDataToChartTime(string knifeCode)
		{

			DataTable dt = TextUtilsHP.LoadDataFromSP("spGetChartProductHistoryData", "A", new string[] { "@knifeCode" }, new object[] { knifeCode });
			DataTable dt1 = TextUtilsHP.LoadDataFromSP("spGetChartProductTotalHistory", "A", new string[] { }, new object[] { });
			DataTable dt2 = new DataTable();

			chartControl1.Series[0].DataSource = dt;
			chartControl1.Series[0].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[0].ArgumentDataMember = "CDay";
			chartControl1.Series[0].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "Quantity" });

			chartControl1.Series[2].DataSource = dt2;


			chartControl1.Series[1].DataSource = dt1;
			chartControl1.Series[1].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[1].ArgumentDataMember = "CDay";
			chartControl1.Series[1].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[1].ValueDataMembers.AddRange(new string[] { "Quantity" });
			//});
		}
		#endregion

		#region Events
		private void timer1_Tick(object sender, EventArgs e)
		{
			if (_refreshFlag == 0)
			{
				_refreshFlag = 1;
				LoadDataToChartProduct();
			}
		}

		private void frmKnifeProcessedChart_Load(object sender, EventArgs e)
		{
			//LoadDataToChartProduct();
			timer1.Enabled = true;
			// Thuoc tinh RuntimeHitTesting can phai duoc bat thi bieu do khi click vao moi co the cho ra du lieu!!!
			//chartControl2.RuntimeHitTesting = true;
		}

		private void chartControl2_CustomDrawSeriesPoint(object sender, CustomDrawSeriesPointEventArgs e)
		{
			if (!string.IsNullOrEmpty(_columnName))
			{
				if (string.Equals(e.SeriesPoint.Argument, _columnName))
				{
					e.SeriesDrawOptions.Color = Color.Red;
					//e.SeriesDrawOptions.ActualColor2 = Color.Red;
				}
			}
		}

		private void frmKnifeProcessedChart_FormClosing(object sender, FormClosingEventArgs e)
		{
			_exitFlag = 2;
			Thread.Sleep(5);
		}
		#endregion

	}
}
