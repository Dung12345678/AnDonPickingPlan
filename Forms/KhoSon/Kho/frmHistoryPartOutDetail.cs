using ST.Business;
using ST.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Media;

namespace BMS
{
	public partial class frmHistoryPartOutDetail : _Forms
	{
		public int _PartOutID;
		public string _OrderCode;
		public string _Pid;
		public string _StockCode;
		public frmHistoryPartOutDetail()
		{
			InitializeComponent();
		}

		private void frmHistoryCheckDetail_Load(object sender, EventArgs e)
		{
			this.Text = _OrderCode + "_" + _Pid + "_" + _StockCode;
			loadData();
		}
		void loadData()
		{
			DataTable dt = TextUtilsStock.LoadDataFromSP("spLoadPartOutDetial", "A"
				, new string[] { "PartOutID" }
				, new object[] { _PartOutID });
			grdData.DataSource = dt;
		}
	}
}
