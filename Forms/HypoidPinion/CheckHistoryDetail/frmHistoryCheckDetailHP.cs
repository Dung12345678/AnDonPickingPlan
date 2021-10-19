using IE.Business;
using IE.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMS
{
	public partial class frmHistoryCheckDetailHP : _Forms
	{
		public DateTime _DateLR;
		public int CheckHistoryID;
		public int Check; //0:Dao,1:Jig,2:Khuôn
		public frmHistoryCheckDetailHP()
		{
			InitializeComponent();
		}
		private void frmHistoryCheckOrderDetailH_Load(object sender, EventArgs e)
		{
			loadData();
			if (Check == 1)
			{
			
				colRealValue2.Visible = true;
				colRealValue2.VisibleIndex = 5;
				colRealValue3.Visible = true;
				colRealValue3.VisibleIndex = 6;
				colRealValue4.Visible = true;
				colRealValue4.VisibleIndex = 7;
				colRealValue5.Visible = true;
				colRealValue5.VisibleIndex = 8;
			}
			else
			{
				colRealValue2.Visible = false;
				colRealValue5.Visible = false;
				colRealValue3.Visible = false;
				colRealValue4.Visible = false;
			}
		}
		void loadData()
		{
			DataTable dt = TextUtilsHP.LoadDataFromSP("spLoadCheckHistoryDetail", "A"
				, new string[] { "@CheckHistoryID" }
				, new object[] { CheckHistoryID });
			//int count = dt.Columns.Count;
			//if (count > 9)
			//{
			//	string RealValue1 = Lib.ToString(dt.Columns[9]);
			//	colRealValue1.FieldName = RealValue1;
			//	colRealValue1.Caption = RealValue1;
			//}
			//else
			//	colRealValue1.Visible = false;
			//if (count > 10)
			//{
			//	string realvalue2 = Lib.ToString(dt.Columns[10]);
			//	colRealValue2.FieldName = realvalue2;
			//	colRealValue2.Caption = realvalue2;
			//}
			//else
			//	colRealValue2.Visible = false;
			//if (count > 11)
			//{
			//	string realvalue3 = Lib.ToString(dt.Columns[11]);
			//	colRealValue3.FieldName = realvalue3;
			//	colRealValue3.Caption = realvalue3;
			//}
			//else
			//	colRealValue3.Visible = false;
			//if (count > 12)
			//{
			//	string realvalue4 = Lib.ToString(dt.Columns[12]);
			//	colRealValue4.FieldName = realvalue4;
			//	colRealValue4.Caption = realvalue4;
			//}
			//else
			//	colRealValue4.Visible = false;
			//if (count > 13)
			//{
			//	string realvalue5 = Lib.ToString(dt.Columns[13]);
			//	colRealValue5.FieldName = realvalue5;
			//	colRealValue5.Caption = realvalue5;
			//}
			//else
			//	colRealValue5.Visible = false;
			grdData.DataSource = dt;
		}
		private void btnExportExecl_Click(object sender, EventArgs e)
		{

		}
	}
}
