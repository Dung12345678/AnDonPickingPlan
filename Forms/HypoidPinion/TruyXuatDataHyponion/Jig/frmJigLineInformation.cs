using DevExpress.XtraGrid.Views.Grid.ViewInfo;
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
	public partial class frmJigLineInformation : _Forms
	{
		#region Variables

		public string _JigCode = "";
		public string _Order = "";
		public string _Goods = "";
		#endregion Variables
		public frmJigLineInformation()
		{
			InitializeComponent();
		}

		private void frmDaoJigLine_Load(object sender, EventArgs e)
		{
			DataSet ds = TextUtilsHP.LoadDataSetFromSP("spGetCheckHistoryJig", new string[] { "@Goods", "@Order", "@KnifeCode" }, new object[] { _Goods, _Order, _JigCode });
			DataTable dtCheckHistory = ds.Tables[0];
			DataTable dtCheckHistoryDetails = ds.Tables[1];
			if (dtCheckHistory.Rows.Count > 0)
			{
				txtGoods.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["GoodsCode"]);
				txtOrder.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["OrderCode"]);
				txtWorker.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["WorkerCode"]);
				txtJigCode.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["KnifeDetailCode"]);
				cbPart.SelectedValue = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["PartID"]);
				cboStep.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["StepCode"]);
				txtDateTime.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["CreateDate"]);
				txtMachine.Text = TextUtilsHP.ToString(dtCheckHistory.Rows[0]["Machine"]);
			}
			if (dtCheckHistoryDetails.Rows.Count > 0)
			{
				grdData.DataSource = dtCheckHistoryDetails;
			}
		}

		private void grvData_Click(object sender, EventArgs e)
		{
			LoadDataToChartTime();
		}
		void LoadDataToChartTime()
		{

			DataTable dt = TextUtilsHP.GetDataTableFromSP("spGetChartProductHistoryJig", new string[] { "@KnifeDetailCode", "@SortOrder" }, new object[] { _JigCode, TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSortOrder)) });
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
			chartControl1.Series[0].ArgumentDataMember = "CDay";
			chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "Quantity" });

			chartControl1.Series[1].DataSource = table;
			chartControl1.Series[1].ArgumentDataMember = "CDay";
			chartControl1.Series[1].ValueDataMembers.AddRange(new string[] { "Max" });

			chartControl1.Series[2].DataSource = table;
			chartControl1.Series[2].ArgumentDataMember = "CDay";
			chartControl1.Series[2].ValueDataMembers.AddRange(new string[] { "Min" });
		}
	}
}
