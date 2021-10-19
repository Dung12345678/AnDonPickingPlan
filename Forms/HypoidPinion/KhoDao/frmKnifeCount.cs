using HP;
using HP.Business;
using HP.Model;
using HP.Utils;
using DevExpress.Data.Filtering.Helpers;
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

namespace BMS
{
	public partial class frmKnifeCount : _Forms
	{
		public frmKnifeCount()
		{
			InitializeComponent();
		}
		private void frmProceductionPlan_Load(object sender, EventArgs e)
		{
			dtpFrom.Value = DateTime.Now;//.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = DateTime.Now;//.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			LoadInfoSearch();
		}


		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportKnifeCount frm = new frmImportKnifeCount();
			frm.Show();
		}
		void LoadInfoSearch()
		{
			DataTable dt = new DataTable();
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			dt = TextUtilsHP.LoadDataFromSP(
					   "spLoadKinfeCount"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@Text" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss"),
										dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss"),
										txtAssyOrderId.Text.Trim()
					   });

			grdData.DataSource = dt;
		}

		private void btnFindDate_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}
	}
}
