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
	public partial class frmJigChart : _Forms
	{
		#region Variables
		public string _knifeCode = "";
		public int _SortOrder = 0;
		#endregion

		#region Methods
		public frmJigChart()
		{
			InitializeComponent();
		}
		#endregion

		#region Events
		private void frmKnifeProcessedChart_Load(object sender, EventArgs e)
		{
			LoadDataToChartTime();
		}
		void LoadDataToChartTime()
		{

			DataTable dt = TextUtilsHP.GetDataTableFromSP("spGetChartProductHistoryJig", new string[] { "@KnifeDetailCode", "@SortOrder" }, new object[] { _knifeCode, _SortOrder });
			//DataTable dtReal = ds.Tables[0];
			//DataTable dtMax = ds.Tables[1];
			//DataTable dtMin = ds.Tables[2];
			DataTable table = new DataTable();
			DataRow row = null;
			table.Columns.Add("CDay", typeof(string));
			table.Columns.Add("Quantity", typeof(decimal));
			table.Columns.Add("Max", typeof(decimal));
			table.Columns.Add("Min", typeof(decimal));

			for (int i = 0; i < dt.Rows.Count; i++)
			{
				row = table.NewRow();
				row["CDay"] = TextUtilsHP.ToString(dt.Rows[i]["CDay"]);
				row["Quantity"] = TextUtilsHP.ToDecimal(dt.Rows[i]["RealValue"]);
				row["Max"] = TextUtilsHP.ToDecimal(dt.Rows[i]["Max"]);
				row["Min"] = TextUtilsHP.ToDecimal(dt.Rows[i]["Min"]);
				table.Rows.Add(row);
			}

			chartControl1.Series[0].DataSource = table;
			chartControl1.Series[0].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[0].ArgumentDataMember = "CDay";
			chartControl1.Series[0].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "Quantity" });

			chartControl1.Series[1].DataSource = table;
			chartControl1.Series[1].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[1].ArgumentDataMember = "CDay";
			chartControl1.Series[1].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[1].ValueDataMembers.AddRange(new string[] { "Max" });

			chartControl1.Series[2].DataSource = table;
			chartControl1.Series[2].ArgumentScaleType = ScaleType.Auto;
			chartControl1.Series[2].ArgumentDataMember = "CDay";
			chartControl1.Series[2].ValueScaleType = ScaleType.Numerical;
			chartControl1.Series[2].ValueDataMembers.AddRange(new string[] { "Min" });
		}
		#endregion
	}
}
