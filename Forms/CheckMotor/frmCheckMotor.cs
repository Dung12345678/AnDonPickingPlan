using SD.Business;
using SD.Model;
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
	public partial class frmCheckMotor : _Forms
	{
		int prevRow;
		public frmCheckMotor()
		{
			InitializeComponent();
		}
		private void frmCheckMotor_Load(object sender, EventArgs e)
		{
			LoadOrderPart();
		}
		#region Method
		void LoadOrderPart()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			DataTable dataTable = TextUtilsSD.LoadDataFromSP(
					   "spLoadCheckMotor"
					   , "A"
					   , new string[] { "@DateStart", "@DateEnd ", "@Text" }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, txbSearch.Text.Trim()
				   });
			grdData.DataSource = dataTable;
		}
		/// <summary>
		/// All form events
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
	
		private void btnCreatePlan_Click(object sender, EventArgs e)
		{
			frmCheckMotorDetails frm = new frmCheckMotorDetails(1);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadOrderPart();
			}
		}

		private void btnEditPlan_Click(object sender, EventArgs e)
		{
			if (grvData.GetSelectedRows().Length != 0)
				prevRow = grvData.GetSelectedRows()[0];
			DataRow row = grvData.GetFocusedDataRow();
			if (row != null)
			{
				string strID = TextUtilsSD.ToString(grvData.GetFocusedRowCellValue(colCardNo));
				CheckMotorModel model = (CheckMotorModel)CheckMotorBO.Instance.FindByAttribute("CardNo", strID)[0];
				frmCheckMotorDetails frm = new frmCheckMotorDetails(2);
				frm._CheckMotorModel = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					LoadOrderPart();
					//  Tu dong focus lai ve dong vua chon
					if (grvData.GetSelectedRows().Length != 0)
						grvData.FocusedRowHandle = prevRow;
				}
			}
		}

		private void btnDelPlan_Click(object sender, EventArgs e)
		{
			if (!grvData.IsDataRow(grvData.FocusedRowHandle))
				return;
			int strID = TextUtilsSD.ToInt(grvData.GetFocusedRowCellValue("ID"));
			string str = TextUtilsSD.ToString(grvData.GetFocusedRowCellValue("PartCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa kế hoạch [{0}] không?", str), TextUtilsSD.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					CheckMotorBO.Instance.Delete(strID);
					LoadOrderPart();
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtilsSD.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void gvSonPlan_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0 && btnEditPlan.Enabled == true)
			{
				btnEditPlan_Click(null, null);
			}
		}

		private void btnSearch_Click(object sender, EventArgs e)
		{
			LoadOrderPart();
		}
		#endregion

		private void dtpFrom_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearch_Click(null, null);
			}
		}

		private void dtpTo_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
			{
				btnSearch_Click(null, null);
			}
		}

		private void txbSearch_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			btnSearch_Click(null, null);
		}
	}
}
