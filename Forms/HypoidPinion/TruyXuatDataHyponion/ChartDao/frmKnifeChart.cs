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
	public partial class frmKnifeChart : _Forms
	{
		#region Variables
		public string _knifeCode = "";
		#endregion

		#region Methods
		public frmKnifeChart()
		{
			InitializeComponent();
		}
		#endregion

		#region Events
		private void frmKnifeProcessedChart_Load(object sender, EventArgs e)
		{
			LoadDataToChart();
		}
		void LoadDataToChart()
		{
			DataSet ds = TextUtilsHP.GetListDataFromSP("spGetChartProductHistoryData", new string[] { "@knifeCode" }, new object[] { _knifeCode });
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
		#endregion
	}
}
