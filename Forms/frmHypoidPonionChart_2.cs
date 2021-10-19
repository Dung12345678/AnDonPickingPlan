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

namespace BMS
{
    public partial class frmHypoidPonionChart_2 : _Forms
    {
        #region Variables
        public delegate void FontSize(decimal fontSize1, decimal fontSize2, decimal fontSize3, decimal fontSize4, decimal fontSize5, decimal fontSize6, decimal fontSize7);
        DataSet dtHYP = new DataSet();
        DataSet dtALTAX = new DataSet();
        string ParthExcelHYP = Application.StartupPath + @"\BieudoHYP.xlsx";
        string ParthExcelALTAX = Application.StartupPath + @"\BieudoALTAX.xlsx";

        #endregion

        #region Method
        public frmHypoidPonionChart_2()
        {
            InitializeComponent();
        }

        #region Event
        private void frmAreaDepartHyp_Load(object sender, EventArgs e)
        {
            dtpTo.Value = DateTime.Now;//.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
            dtpFrom.Value = DateTime.Now.AddDays(-30);//.Date.AddHours(0).AddMinutes(00).AddSeconds(00);

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
        #endregion


        #endregion

        //private void grvHYP_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        //{
        //try
        //{
        //    string Statuss = Lib.ToString(grvHYP.GetRowCellValue(e.RowHandle, colStatusHYP));
        //    string[] name = Statuss.Split(' ');
        //    if (Statuss.ToUpper().Contains("ĐÃ XONG"))
        //    {
        //        e.Appearance.BackColor = Color.Lime;
        //    }
        //    else
        //    {
        //        if (name[2].ToString() == "1" || name[2].ToString() == "2" || name[2].ToString() == "3")
        //        {
        //            e.Appearance.BackColor = Color.Yellow;
        //        }
        //        else
        //        {
        //            e.Appearance.BackColor = Color.Red;

        //        }
        //    }
        //}
        //catch (Exception)
        //{

        //}
        //}

        /// <summary>
        /// hiển thị lên biểu đồ HYP
        /// </summary>
        void loadChartHYP()
        {
            try
            {

                spreadsheetControlHYP.BeginUpdate();
                //View biểu đồ 
                spreadsheetControlHYP.LoadDocument(ParthExcelHYP);
                IWorkbook workbook = spreadsheetControlHYP.Document;

                //Biểu đồ tồn lâu HYP
                Worksheet workDay = workbook.Worksheets["DataBieuDoTonLau"];
                //Xóa các dữ liệu có trong bảng 
                CellRange TempRangeDay = workDay.Range["B2:C4"];
                TempRangeDay.Clear();
                //Hiển thị cột 
                workDay.Cells[1, 1].Value = Lib.ToInt(dtHYP.Tables[1].Rows[0][0].ToString());
                workDay.Cells[2, 1].Value = Lib.ToInt(dtHYP.Tables[2].Rows[0][0].ToString());


                //Biểu đồ tồn công đoạn HYP
                Worksheet workMonth = workbook.Worksheets["DataBieuDoTonCongDoan"];
                //Xóa các dữ liệu có trong bảng 
                CellRange TempRangeMonth = workMonth.Range["B2:C5"];
                TempRangeMonth.Clear();
                // hiển thị lên cột
                workMonth.Cells[1, 1].Value = Lib.ToInt(dtHYP.Tables[3].Rows[0][0].ToString());
                workMonth.Cells[2, 1].Value = Lib.ToInt(dtHYP.Tables[4].Rows[0][0].ToString());
                workMonth.Cells[3, 1].Value = Lib.ToInt(dtHYP.Tables[5].Rows[0][0].ToString());

                spreadsheetControlHYP.ReadOnly = false;
                spreadsheetControlHYP.Enabled = false;

            }
            catch (Exception)
            {

            }
            finally
            {
                spreadsheetControlHYP.EndUpdate();
            }
        }

        /// <summary>
        /// hiển thị lên biểu đồ ALTAX
        /// </summary>
        void loadChartALTAX()
        {
            try
            {
                spreadsheetControlALTAX.BeginUpdate();
                //View biểu đồ
                spreadsheetControlALTAX.LoadDocument(ParthExcelALTAX);
                IWorkbook workbook = spreadsheetControlALTAX.Document;

                //Biểu đồ TỒN LÂU ALTAX
                Worksheet workDay = workbook.Worksheets["DataBieuDoTonLau"];
                //Xóa các dữ liệu có trong bảng 
                CellRange TempRangeDay = workDay.Range["B2:C4"];
                TempRangeDay.Clear();
                //Hiển thị cột 
                workDay.Cells[1, 1].Value = TextUtilsHP.ToInt(dtALTAX.Tables[1].Rows[0][0].ToString());
                workDay.Cells[2, 1].Value = TextUtilsHP.ToInt(dtALTAX.Tables[2].Rows[0][0].ToString());

                //Biểu đồ TỒN CÔNG ĐOẠN ALTAX 
                Worksheet workMonth = workbook.Worksheets["DataBieuDoTonCongDoan"];
                //Xóa các dữ liệu có trong bảng 
                CellRange TempRangeMonth = workMonth.Range["B2:C5"];
                TempRangeMonth.Clear();
                // hiển thị lên cột
                workMonth.Cells[1, 1].Value = TextUtilsHP.ToInt(dtALTAX.Tables[3].Rows[0][0].ToString());
                workMonth.Cells[2, 1].Value = TextUtilsHP.ToInt(dtALTAX.Tables[4].Rows[0][0].ToString());
                workMonth.Cells[3, 1].Value = TextUtilsHP.ToInt(dtALTAX.Tables[5].Rows[0][0].ToString());

                spreadsheetControlALTAX.ReadOnly = false;
                spreadsheetControlALTAX.Enabled = false;
            }
            catch (Exception)
            {

            }
            finally
            {
                spreadsheetControlALTAX.EndUpdate();
            }
        }

        private void btnFindDate_Click(object sender, EventArgs e)
        {
            loadHYP();
            loadALTAX();
            loadChartHYP();
            loadChartALTAX();
        }

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
        
        private void btnSettings_Click(object sender, EventArgs e)
        {
            frmConfig frm = new frmConfig();
            frm.Show();
        }
    }
}
