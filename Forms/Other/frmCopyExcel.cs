using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Collections;
using System.Diagnostics;
using BMS.Business;
using BMS.Model;
using BMS.Utils;
using DevExpress.XtraGrid.Views.Base;
using System.Threading;

namespace BMS
{
    public partial class frmCopyExcel : _Forms
    {
        public DataTable source = null;   
        public string PriceDataCategoryIDs;

        public frmCopyExcel()
        {
            InitializeComponent();
        }

        private void frmPriceData_Import_Load(object sender, EventArgs e)
        {
           
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
            this.DialogResult = DialogResult.OK;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < grvData.RowCount; i++)
            {
                string departmentCode = grvData.GetRowCellValue(i, "colPB").ToString();
                if (departmentCode == "TK") continue;

                ProcessTransaction pt = new ProcessTransaction();
                pt.OpenConnection();
                pt.BeginTransaction();
                try
                {
                    UsersModel Model = new UsersModel();
                    try
                    {
                        Model.BirthOfDate = TextUtils.ToDate(grvData.GetRowCellValue(i, "colBirdDate").ToString());
                    }
                    catch
                    {
                        Model.BirthOfDate = DateTime.Now;
                    }

                    Model.CMTND = grvData.GetRowCellValue(i, "colCMT").ToString();
                    Model.Code = grvData.GetRowCellValue(i, "colCode").ToString();
                    
                    Model.DepartmentID = ((DepartmentModel)DepartmentBO.Instance.FindByAttribute("Code", departmentCode)[0]).ID;
                    Model.FullName = grvData.GetRowCellValue(i, "colName").ToString();
                    Model.Qualifications = grvData.GetRowCellValue(i, "colTrinhDo").ToString();
                    try
                    {
                        Model.StartWorking = TextUtils.ToDate(grvData.GetRowCellValue(i, "colStartDate").ToString());
                    }
                    catch
                    {
                        Model.StartWorking = DateTime.Now;
                    }

                    Model.Status = 0;

                    if (grvData.GetRowCellValue(i, "colLogin").ToString()!="")
                    {
                        Model.LoginName = grvData.GetRowCellValue(i, "colLogin").ToString();
                        Model.PasswordHash = MD5.EncryptPassword("123456");
                    }
                    else
                    {
                        Model.LoginName = "";
                        Model.PasswordHash = "";
                    }

                    if (Model.ID == 0)
                    {
                        Model.CreatedDate = TextUtils.GetSystemDate();
                        Model.CreatedBy = Global.AppUserName;
                        Model.UpdatedDate = Model.CreatedDate;
                        Model.UpdatedBy = Global.AppUserName;
                        Model.ID = (int)pt.Insert(Model);
                    }
                    else
                    {
                        Model.UpdatedDate = TextUtils.GetSystemDate();
                        Model.UpdatedBy = Global.AppUserName;
                        pt.Update(Model);
                    }
                    pt.CommitTransaction();                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    pt.CloseConnection();
                }   
            }
        }

        private void grdData_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Control && e.KeyCode == Keys.V)
            {
                grdData.DataSource = null;  

                string[] cb = Clipboard.GetText(TextDataFormat.UnicodeText).Split(new string[1] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                cb[0] = cb[0];
                string[] ColHeaders = cb[0].Split('\t');

                DataTable dTab = new DataTable();
                int colIndex = 1;
                for (int i = 0; i < ColHeaders.Length; i++)
                {
                    if (i == 0)
                    {
                        dTab.Columns.Add("colSTT", Type.GetType("System.String"));
                        dTab.Columns[0].Caption = "STT";
                    }
                    else if (i == 1)
                    {
                        dTab.Columns.Add("colCode", Type.GetType("System.String"));
                        dTab.Columns[1].Caption = "Mã nhân viên";
                    }
                    else if (i == 2)
                    {
                        dTab.Columns.Add("colPB", Type.GetType("System.String"));
                        dTab.Columns[2].Caption = "Phòng ban";
                    }
                    else if (i == 3)
                    {
                        dTab.Columns.Add("colGroup", Type.GetType("System.String"));
                        dTab.Columns[3].Caption = "Nhóm";
                    }
                    else if (i == 4)
                    {
                        dTab.Columns.Add("colLogin", Type.GetType("System.String"));
                        dTab.Columns[4].Caption = "Tên đăng nhập";
                    }
                  
                    else if (i == 5)
                    {
                        dTab.Columns.Add("colPQ", Type.GetType("System.String"));
                        dTab.Columns[5].Caption = "Quyền";
                    }
                    else if (i == 6)
                    {
                        dTab.Columns.Add("colName", Type.GetType("System.String"));
                        dTab.Columns[6].Caption = "Tên NV";
                    }
                    else if (i == 7)
                    {
                        dTab.Columns.Add("colBirdDate", Type.GetType("System.String"));
                        dTab.Columns[7].Caption = "Ngày sinh";
                    }
                    else if (i == 8)
                    {
                        dTab.Columns.Add("colTrinhDo", Type.GetType("System.String"));
                        dTab.Columns[8].Caption = "Trình độ";
                    }
                    else if (i == 9)
                    {
                        dTab.Columns.Add("colCMT", Type.GetType("System.String"));
                        dTab.Columns[9].Caption = "CMTND";
                    }
                    else if (i == 10)
                    {
                        dTab.Columns.Add("colStartDate", Type.GetType("System.String"));
                        dTab.Columns[10].Caption = "Ngày vào CT";
                    }  
                    else if (i == 11)
                    {
                        dTab.Columns.Add("colDescription", Type.GetType("System.String"));
                        dTab.Columns[11].Caption = "Ghi chú";
                    }
                    else if (i == 12)
                    {
                        dTab.Columns.Add("colStartDate", Type.GetType("System.String"));
                        dTab.Columns[12].Caption = "Người duyệt";
                    }
                    else if (i == 13)
                    {
                        dTab.Columns.Add("colFile", Type.GetType("System.String"));
                        dTab.Columns[13].Caption = "File";
                    }
                    else if (i == 14)
                    {
                        dTab.Columns.Add("colCompanyProfile", Type.GetType("System.String"));
                        dTab.Columns[14].Caption = "Tên nhà thầu";
                    }                    
                    else if (i == 15)
                    {
                        dTab.Columns.Add("colCreateBy", Type.GetType("System.String"));
                        dTab.Columns[15].Caption = "Người nhập dữ liệu";
                    }
                    else if (i == 16)
                    {
                        dTab.Columns.Add("colTeamLeader", Type.GetType("System.String"));
                        dTab.Columns[16].Caption = "Trưởng Nhóm KSG";
                    }
                    else if (i == 17)
                    {
                        dTab.Columns.Add("colManager", Type.GetType("System.String"));
                        dTab.Columns[17].Caption = "Trưởng phòng đấu thầu";
                    }
                    else
                    {
                        dTab.Columns.Add("T" + colIndex, Type.GetType("System.String"));
                        colIndex++;
                    }
                }

                DataRow dr = null;
                for (int r = 0; r < cb.Length; r++)
                {
                    if (cb[r].Trim().Length > 0)
                    {
                        dr = dTab.NewRow();
                        string[] row = Convert.ToString(cb[r]).Split('\t');

                        for (int c = 0; c < ColHeaders.Length; c++)
                        {
                            dr[c] = row[c];
                        }

                        dTab.Rows.Add(dr);
                    }
                }
                //foreach (string HeaderStr in ColHeaders)
                //{
                //    dTab.Columns.Add(HeaderStr);
                //}

                grdData.DataSource = dTab;
                //for (int i = 0; i < grvData.Columns.Count;i++ )
                //{
                //    grvData.Columns[i].OptionsColumn.AllowEdit = false;
                //}
                grvData.BestFitColumns();
            }
        }
       
    }
}
