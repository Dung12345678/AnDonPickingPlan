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
using BMS.Business;
using BMS.Model; //Altax
using IE.Business; //Hyp
using IE.Model;
using DevExpress.Spreadsheet;
using System.Threading;

namespace BMS
{
	public partial class frmHypoidPonionChart : _Forms
	{
		#region Variables
		public delegate void FontSize(decimal fontSize1, decimal fontSize2, decimal fontSize3, decimal fontSize4, decimal fontSize5, decimal fontSize6, decimal fontSize7);
		DataSet dtHYP = new DataSet();
		DataSet dtALTAX = new DataSet();
		string ParthExcelHYP = Application.StartupPath + @"\BieudoHYP.xlsx";
		string ParthExcelALTAX = Application.StartupPath + @"\BieudoALTAX.xlsx";
		List<int> lstMaster = new List<int>();
		List<int> lstDetails = new List<int>();
		int Max = 0;
		int MaxDetails = 0;
		string name1 = "Biểu đồ tồn lâu";
		Thread _threadPlan;
		#endregion

		public frmHypoidPonionChart()
		{
			InitializeComponent();

		}

		#region Event
		private void frmAreaDepartHyp_Load(object sender, EventArgs e)
		{
			dtpTo.Value = DateTime.Now;//.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			dtpFrom.Value = DateTime.Now.AddDays(-30);//.Date.AddHours(0).AddMinutes(00).AddSeconds(00);

			_threadPlan = new Thread(new ThreadStart(Plan));
			_threadPlan.IsBackground = true;
			_threadPlan.Start();

			loadHYP();
			loadALTAX();
			loadChartHYP();
			loadChartALTAX();
		}

		/// <summary>
		/// hiển thị danh sách HYP
		/// </summary>
		void loadHYP()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			dtHYP = LibIE.LoadDataSetFromSP(
					   "spGetProductionPlanChart"
					   , new string[] { "@DateStart", "@DateEnd " }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss") });
			grdHYP.DataSource = dtHYP.Tables[0];
		}

		/// <summary>
		/// hiển thị danh sách ALTAX
		/// </summary>
		void loadALTAX()
		{
			dtpFrom.Value = dtpFrom.Value.Date.AddHours(0).AddMinutes(00).AddSeconds(00);
			dtpTo.Value = dtpTo.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
			dtALTAX = TextUtils.LoadDataSetFromSP(
					   "spGetProductionPlanChart"
					   , new string[] { "@DateStart", "@DateEnd " }
					   , new object[] { dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
										, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss") });
			grdALTAX.DataSource = dtALTAX.Tables[0];
		}
		/// <summary>
		/// hiển thị lên HYP
		/// </summary>
		void loadChartHYP()
		{
			try
			{
				chartHYPTonLau.Legends.Clear(); // xóa đi phụ đề  bên cạnh biểu đồ
				chartHYPTonLau.Series["TonLau"].Points.Clear(); // clear đi nhứng cột đã có

				chartHYPTonLau.Series["TonLau"].Points.AddXY("Chưa xong", Lib.ToInt(dtHYP.Tables[6].Rows[0][0].ToString()));
				chartHYPTonLau.Series["TonLau"].Points[0].Color = Color.DarkCyan;
				chartHYPTonLau.Series["TonLau"].Points[0].Label = Lib.ToInt(dtHYP.Tables[6].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtHYP.Tables[6].Rows[0][0]));

				chartHYPTonLau.Series["TonLau"].Points.AddXY("1 ~ 3 ngày", Lib.ToInt(dtHYP.Tables[1].Rows[0][0].ToString()));
				chartHYPTonLau.Series["TonLau"].Points[1].Color = Color.Yellow;
				chartHYPTonLau.Series["TonLau"].Points[1].Label = Lib.ToInt(dtHYP.Tables[1].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtHYP.Tables[1].Rows[0][0]));

				chartHYPTonLau.Series["TonLau"].Points.AddXY("> 3 ngày", Lib.ToInt(dtHYP.Tables[2].Rows[0][0].ToString()));
				chartHYPTonLau.Series["TonLau"].Points[2].Color = Color.Red;
				chartHYPTonLau.Series["TonLau"].Points[2].Label = Lib.ToInt(dtHYP.Tables[2].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtHYP.Tables[2].Rows[0][0]));

				chartHYPTonLau.Update();


				chartHYPCongDoan.Legends.Clear(); // XÓA BỎ TÊN BÊN CẠNH biểu đồ

				chartHYPCongDoan.Series["CongDoan"].Points.Clear(); // clear đi nhứng cột đã có
				chartHYPCongDoan.Series["CongDoan"].Points.AddXY("KHO", Lib.ToInt(dtHYP.Tables[3].Rows[0][0].ToString()));
				chartHYPCongDoan.Series["CongDoan"].Points[0].Label = TextUtilsHP.ToInt(dtHYP.Tables[3].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtHYP.Tables[3].Rows[0][0]));

				chartHYPCongDoan.Series["CongDoan"].Points.AddXY("ASSY", Lib.ToInt(dtHYP.Tables[4].Rows[0][0].ToString()));
				chartHYPCongDoan.Series["CongDoan"].Points[1].Label = Lib.ToInt(dtHYP.Tables[4].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtHYP.Tables[4].Rows[0][0]));

				chartHYPCongDoan.Series["CongDoan"].Points.AddXY("QI", Lib.ToInt(dtHYP.Tables[5].Rows[0][0].ToString()));
				chartHYPCongDoan.Series["CongDoan"].Points[2].Label = Lib.ToInt(dtHYP.Tables[5].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtHYP.Tables[5].Rows[0][0]));

				chartHYPCongDoan.Series["CongDoan"].Points.AddXY("Other", Lib.ToInt(dtHYP.Tables[7].Rows[0][0].ToString()));
				chartHYPCongDoan.Series["CongDoan"].Points[3].Label = Lib.ToInt(dtHYP.Tables[7].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtHYP.Tables[7].Rows[0][0]));

				chartHYPCongDoan.Update();

			}
			catch (Exception)
			{

			}
		}

		/// <summary>
		/// hiển thị lên biểu đồ ALTAX
		/// </summary>
		void loadChartALTAX()
		{
			try
			{
				chartALTAXTonLau.Legends.Clear(); // XÓA BỎ TÊN BÊN CẠNH biểu đồ
				chartALTAXTonLau.Series["TonLau"].Points.Clear(); // clear đi nhứng cột đã có

				chartALTAXTonLau.Series["TonLau"].Points.AddXY("Chưa xong", TextUtilsHP.ToInt(dtALTAX.Tables[6].Rows[0][0].ToString()));
				chartALTAXTonLau.Series["TonLau"].Points[0].Color = Color.DarkCyan;
				chartALTAXTonLau.Series["TonLau"].Points[0].Label = TextUtilsHP.ToInt(dtALTAX.Tables[6].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtALTAX.Tables[6].Rows[0][0]));

				chartALTAXTonLau.Series["TonLau"].Points.AddXY("1 ~ 3 ngày", TextUtilsHP.ToInt(dtALTAX.Tables[1].Rows[0][0].ToString()));
				chartALTAXTonLau.Series["TonLau"].Points[1].Color = Color.Yellow;
				chartALTAXTonLau.Series["TonLau"].Points[1].Label = TextUtilsHP.ToInt(dtALTAX.Tables[1].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtALTAX.Tables[1].Rows[0][0]));

				chartALTAXTonLau.Series["TonLau"].Points.AddXY("> 3 ngày", TextUtilsHP.ToInt(dtALTAX.Tables[2].Rows[0][0].ToString()));
				chartALTAXTonLau.Series["TonLau"].Points[2].Color = Color.Red;
				chartALTAXTonLau.Series["TonLau"].Points[2].Label = TextUtilsHP.ToInt(dtALTAX.Tables[2].Rows[0][0]).ToString();
				lstMaster.Add(Lib.ToInt(dtALTAX.Tables[2].Rows[0][0]));

				chartALTAXTonLau.Update();

				//
				chartALTAXCongDoan.Legends.Clear(); // XÓA BỎ TÊN BÊN CẠNH biểu đồ

				chartALTAXCongDoan.Series["CongDoan"].Points.Clear(); // clear đi nhứng cột đã có
				chartALTAXCongDoan.Series["CongDoan"].Points.AddXY("KHO", TextUtilsHP.ToInt(dtALTAX.Tables[3].Rows[0][0].ToString()));
				chartALTAXCongDoan.Series["CongDoan"].Points[0].Label = TextUtilsHP.ToInt(dtALTAX.Tables[3].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtALTAX.Tables[3].Rows[0][0]));

				chartALTAXCongDoan.Series["CongDoan"].Points.AddXY("ASSY", TextUtilsHP.ToInt(dtALTAX.Tables[4].Rows[0][0].ToString()));
				chartALTAXCongDoan.Series["CongDoan"].Points[1].Label = TextUtilsHP.ToInt(dtALTAX.Tables[4].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtALTAX.Tables[4].Rows[0][0]));

				chartALTAXCongDoan.Series["CongDoan"].Points.AddXY("QI", TextUtilsHP.ToInt(dtALTAX.Tables[5].Rows[0][0].ToString()));
				chartALTAXCongDoan.Series["CongDoan"].Points[2].Label = TextUtilsHP.ToInt(dtALTAX.Tables[5].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtALTAX.Tables[5].Rows[0][0]));

				chartALTAXCongDoan.Series["CongDoan"].Points.AddXY("Other", TextUtilsHP.ToInt(dtALTAX.Tables[7].Rows[0][0].ToString()));
				chartALTAXCongDoan.Series["CongDoan"].Points[3].Label = TextUtilsHP.ToInt(dtALTAX.Tables[7].Rows[0][0]).ToString();
				lstDetails.Add(TextUtilsHP.ToInt(dtALTAX.Tables[7].Rows[0][0]));

				chartALTAXCongDoan.Update();
			}
			catch (Exception)
			{

			}
		}
		#endregion

		#region Buttons Event
		/// <summary>
		/// click button tìm kiếm
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnFindDate_Click(object sender, EventArgs e)
		{
			loadHYP();
			loadALTAX();
			loadChartHYP();
			loadChartALTAX();
		}
		void Plan()
		{
			while (true)
			{
				Thread.Sleep(120000);
				try
				{
					this.Invoke((MethodInvoker)delegate
					{
						lstMaster.Clear();
						lstDetails.Clear();
						loadHYP();
						loadALTAX();
						loadChartHYP();
						loadChartALTAX();
						int MaxMaster = 200;
						int MaxDetails = 200;
						while (lstMaster.Max() > MaxMaster)
						{
							MaxMaster += 100;
						}
						while (lstDetails.Max() > MaxDetails)
						{
							MaxDetails += 200;
						}

						chartHYPTonLau.ChartAreas[0].Axes[1].Maximum = MaxMaster;
						chartHYPCongDoan.ChartAreas[0].Axes[1].Maximum = MaxDetails;
						chartALTAXTonLau.ChartAreas[0].Axes[1].Maximum = MaxMaster;
						chartALTAXCongDoan.ChartAreas[0].Axes[1].Maximum = MaxDetails;

					});
				}
				catch
				{
				}
			}
		}

		/// <summary>
		/// click button lưu
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSave_Click(object sender, EventArgs e)
		{
			try
			{
				// hyp
				for (int i = 0; i < grvHYP.RowCount; i++)
				{
					int id = Lib.ToInt(grvHYP.GetRowCellValue(i, colIDHYP));
					IE.Model.ProductionPlanModel productionPlanHYP = new IE.Model.ProductionPlanModel();
					if (id > 0)
					{
						productionPlanHYP = (IE.Model.ProductionPlanModel)(IE.Business.ProductionPlanBO.Instance.FindByPK(id));
					}
					productionPlanHYP.ID = id;
					productionPlanHYP.Cause = Lib.ToString(grvHYP.GetRowCellValue(i, colCauseHYP));
					if (productionPlanHYP.ID > 0)
					{
						IE.Business.ProductionPlanBO.Instance.Update(productionPlanHYP);
					}
				}

				// Altax
				for (int i = 0; i < grvALTAX.RowCount; i++)
				{
					int id = TextUtilsHP.ToInt(grvALTAX.GetRowCellValue(i, colIDALTAX));
					BMS.Model.ProductionPlanModel productionPlan = new BMS.Model.ProductionPlanModel();
					if (id > 0)
					{
						productionPlan = (BMS.Model.ProductionPlanModel)(BMS.Business.ProductionPlanBO.Instance.FindByPK(id));
					}
					productionPlan.ID = id;
					productionPlan.Cause = TextUtilsHP.ToString(grvALTAX.GetRowCellValue(i, colCauseALTAX));
					if (productionPlan.ID > 0)
					{
						BMS.Business.ProductionPlanBO.Instance.Update(productionPlan);
					}
				}
			}
			catch (Exception)
			{
				throw;
			}
		}
		#endregion

		private void grvHYP_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
			int value = TextUtils.ToInt(grvHYP.GetRowCellValue(e.RowHandle, colShowColor));
			switch (value)
			{
				case 1:
					e.Appearance.BackColor = Color.IndianRed;
					break;
				case 2:
					e.Appearance.BackColor = Color.Yellow;
					break;
				default:
					break;
			}
		}

		private void grvALTAX_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
			int value = TextUtils.ToInt(grvALTAX.GetRowCellValue(e.RowHandle, colShowColorAltax));
			switch (value)
			{
				case 1:
					e.Appearance.BackColor = Color.IndianRed;
					break;
				case 2:
					e.Appearance.BackColor = Color.Yellow;
					break;
				default:
					break;
			}
		}
	}
}
