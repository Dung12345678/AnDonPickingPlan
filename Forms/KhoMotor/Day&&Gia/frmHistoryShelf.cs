using BMS.Business;
using BMS.Model;
using ST.Business;
using ST.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmHistoryShelf : _Forms
	{
		public frmHistoryShelf()
		{
			InitializeComponent();
		}
		private void frmHistoryPartOut_Load(object sender, EventArgs e)
		{
			LoadData();
		}
		void LoadData()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

			DataTable dataTable = TextUtilsStock.LoadDataFromSP(
						"spLoadHistoryShelfMotor"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@Text" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txbSearchHistory.Text.Trim()
				   });
			grdData.DataSource = dataTable;
		}

		private void btnSearchHistory_Click(object sender, EventArgs e)
		{
			LoadData();
		}
		private void dtpFrom_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearchHistory_Click(null, null);
			}
		}
		private void dtpTo_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearchHistory_Click(null, null);
			}
		}

		private void txbSearchHistory_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			btnSearchHistory_Click(null, null);
		}
	}
}
