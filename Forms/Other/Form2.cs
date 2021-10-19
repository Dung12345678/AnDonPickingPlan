using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using DevExpress.Utils;
using BMS.Model;
using BMS.Business;
using System.Collections;
using BMS.Utils;
using TPA.Model;
using TPA.Business;
using MODI;

namespace BMS
{
    public partial class Form2 : _Forms
    {
        DataTable dtData = null;
        DateTime _pStart = new DateTime(2015, 1, 1);
        DateTime _pEnd = new DateTime(2015, 2, 1);

        List<DateInfo> allDates = new List<DateInfo>();

        public class DateInfo
        {
            public DateTime ThisDate { get; set; }
            public string DayInWeek { get; set; }
        }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtData = new DataTable();
            dtData.Columns.Add("STT");
            dtData.Columns.Add("Name");
            dtData.Columns.Add("Author");
            dtData.Columns.Add("StartDate", typeof(DateTime));
            dtData.Columns.Add("FinishDate", typeof(DateTime));
            dtData.Columns.Add("CellsNumber");
            dtData.Columns.Add("Space");

            dtData.Rows.Add("1", "Chỉnh sửa bản vẽ 3D", "Phòng TK", new DateTime(2015, 1, 1), new DateTime(2015, 1, 8), 8,0);
            dtData.Rows.Add("2", "Canon đặt hàng", "Phòng Dự án", new DateTime(2015, 1, 9), new DateTime(2015, 1, 13), 5,8);
            dtData.Rows.Add("3", "Đặt hàng thiết bị", "Phòng vật tư", new DateTime(2015, 1, 14), new DateTime(2015, 1, 15), 2,13);
            
            //loadTree();
            //loadGrid();

        }

        void loadTree(string projectCode)
        {
            try
            {
                using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang load danh sách module..."))
                {
                    string[] _paraName = new string[1];
                    object[] _paraValue = new object[1];
                    _paraName[0] = "@ProjectCode"; _paraValue[0] = "T009.1501";
                    //DataTable Source = ModulesBO.Instance.LoadDataFromSP("spGetProductOfProjectNew", "Source", _paraName, _paraValue);
                    DataTable Source = LibQLSX.Select("select * from vGetProductOfProject where ProjectCode ='" + projectCode + "'");
                    treeData.DataSource = Source;
                    treeData.KeyFieldName = "PProductId";
                    treeData.PreviewFieldName = "MaThietBi";
                    //treeData.ExpandAll();                
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        void loadGrid()
        {
            DataTable dt = TextUtils.Select("select * from Test");
            gridControl1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //for (DateTime date = _pStart; date <= _pEnd; date = date.AddDays(1))
            //{
            //    DateInfo di = new DateInfo();
            //    di.DayInWeek = date.ToString("ddd");
            //    di.ThisDate = date;
            //    allDates.Add(di);
            //}

            TextUtils.WordFindAndReplace(@"C:\Users\thao.nv\Desktop\PTTK.TPAD.C4802.docm", "TPAD.C4802", "TPAD.C4800");
            MessageBox.Show("ok!");

            //string excelFile = @"C:\Users\thao.nv\Desktop\Test Excel.xlsx";

            //Excel.Application objXLApp = default(Excel.Application);
            //Excel.Workbook objXLWb = default(Excel.Workbook);
            //Excel.Worksheet objXLWs = default(Excel.Worksheet);

            ////File.Copy("D:\\Test Excel.xlsx", "D:\\1.xlsx", true); // Copy file hồ sơ thiết kế

            //objXLApp = new Excel.Application();
            //objXLApp.Workbooks.Open(excelFile);
            //objXLWb = objXLApp.Workbooks[1];
            //objXLWs = (Excel.Worksheet)objXLWb.Worksheets[1];

            ////for (int i = 0; i <= allDates.Count - 1; i++)
            ////{
            ////    objXLWs.Cells[5, 4 + i] = allDates[i].ThisDate.ToString("dd/MM");
            ////    objXLWs.Cells[6, 4 + i] = allDates[i].DayInWeek;
            ////}

            ////objXLWs.Shapes.AddPicture(@"C:\Users\thao.nv\Desktop\Untitled.png", Microsoft.Office.Core.MsoTriState.msoFalse,
            ////    Microsoft.Office.Core.MsoTriState.msoCTrue, 50, 50, 900, 200);

            //objXLWs.Cells.Replace(What: "TPAD.C9904", Replacement: "eee", LookAt: Excel.XlLookAt.xlWhole,
            //SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: true, ReplaceFormat: false);

            //objXLWs.PageSetup.LeftHeader = objXLWs.PageSetup.LeftHeader.Replace("TPAD.C9904", "eee");
            //objXLWs.PageSetup.CenterHeader = objXLWs.PageSetup.CenterHeader.Replace("TPAD.C9904", "eee");
            //objXLWs.PageSetup.RightHeader = objXLWs.PageSetup.RightHeader.Replace("TPAD.C9904", "eee");
            //objXLWs.PageSetup.LeftFooter = objXLWs.PageSetup.LeftFooter.Replace("TPAD.C9904", "eee");
            //objXLWs.PageSetup.CenterFooter = objXLWs.PageSetup.CenterFooter.Replace("TPAD.C9904", "eee");
            //objXLWs.PageSetup.RightFooter = objXLWs.PageSetup.RightFooter.Replace("TPAD.C9904", "eee"); 


            //objXLApp.ActiveWorkbook.Save();
            //objXLApp.Workbooks.Close();
            //objXLApp.Quit();

            //Process.Start(excelFile);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //frmMaterialImportExcel frm = new frmMaterialImportExcel();
            //frm.Show();
            frmShowExcel frm = new frmShowExcel();
            frm.Show();
        }

        private void btnTestGridBackground_Click(object sender, EventArgs e)
        {
            frmBackgroundGrid frm = new frmBackgroundGrid();
            TextUtils.OpenForm(frm);
        }

        private void btnErrorExcel_Click(object sender, EventArgs e)
        {
            frmErrorExcel frm = new frmErrorExcel();
            TextUtils.OpenForm(frm);
        }

        private void btnForm1_Click(object sender, EventArgs e)
        {
            Form1 frm = new Form1();
            TextUtils.OpenForm(frm);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (WaitDialogForm fWait1 = new WaitDialogForm("Vui lòng chờ trong giây lát", "Tạo", new Size(this.Width / 4, this.Height / 9), this))
            {
                DataTable dtModule = TextUtils.Select("select * from Modules with(nolock) where status = 2 and Code like '%tpad%' order by Code");
                foreach (DataRow item in dtModule.Rows)
                {
                    string productCode = item["Code"].ToString();
                    DataTable dt = TextUtils.Select("select * from ModuleVersion with(nolock) where ModuleCode = '" + productCode + "'");
                    if (dt.Rows.Count == 0)
                    {
                        using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát", "Tạo " + productCode, new Size(this.Width / 4, this.Height / 9), this))
                        {
                            try
                            {
                                string path = TextUtils.DownloadAll(productCode);

                                ModuleVersionModel model = new ModuleVersionModel();
                                //model.ProjectCode = misMatchModel.ProjectCode;
                                model.ModuleCode = productCode;
                                //model.MisMatchCode = misMatchModel.Code;
                                model.Version = 0;
                                model.Path = path;
                                model.Description = "Tạo phiên bản đầu tiên của module";
                                model.Reason = "Tạo phiên bản đầu tiên của module";
                                ModuleVersionBO.Instance.Insert(model);

                                //MessageBox.Show("Tạo phiên bản đầu tiên của module [" + productCode + "] thành công!");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Lỗi: " + ex.Message);
                            }
                        }
                    }
                }
            }
        }

        void Test()
        {
            string a = "TPAD";
            string a0 = "TPAD.B4501";
            string a1 = "TPAD.B4501.01";
            string a2 = "TPAD.B4501.01.01";
            string text = a + " - " + a.Split('.').Count() + Environment.NewLine +
                a0 + " - " + a0.Split('.').Count() + Environment.NewLine +
                a1 + " - " + a1.Split('.').Count() + Environment.NewLine +
                a2 + " - " + a2.Split('.').Count();
            MessageBox.Show(text);
        }

        void ChangeModuleName()
        {
            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang lấy mã mới"))
            {
                string tableCodeFilePath = @"\\server\data2\Thietke\ISO\ISO.Thietke\TAI LIEU DAO TAO\TAI LIEU HO TRO PHONG THIET KE\TK09- Huong dan doi ma san pham TK\TK09-BM01 - Bang thay doi ma TK.xlsx";
                List<string> listSheet = TextUtils.ListSheetInExcel(tableCodeFilePath);
                foreach (string sheetName in listSheet)
                {
                    if (!sheetName.ToUpper().StartsWith("TPAD.")) continue;

                    DataTable dt = TextUtils.ExcelToDatatableNoHeader(tableCodeFilePath, sheetName);
                    foreach (DataRow item in dt.Rows)
                    {                        
                        string newCode = item["F3"] == null ? "" : item["F3"].ToString();
                        if (newCode == "") continue;

                        ModulesModel model = null;
                        try
                        {
                            model = (ModulesModel)ModulesBO.Instance.FindByCode(newCode);
                        }
                        catch (Exception)
                        {                          
                        }
                        if (model != null)
                        {
                            string oldName = item["F5"] == null ? "" : item["F5"].ToString();
                            string newName = item["F6"] == null ? "" : item["F6"].ToString();
                            string name = (newName == "" ? oldName : newName);

                            model.Name = name;
                            ModulesBO.Instance.Update(model);
                        }                        
                    }                    
                }
            }
        }

        void DownloadDMVT()
        {
            DocUtils.InitFTPQLSX();
            ArrayList listModules = ModulesBO.Instance.FindByExpression(new Expression("Status", 2).And(new Expression("Code", "TPAD", "like")));
            int count = 0;
            foreach (var item in listModules)
            {
                ModulesModel module = (ModulesModel)item;
                string moduleCode = module.Code;
                string dmvtPath = string.Format("Thietke.Ck/{0}/{1}.Ck/VT.{1}.xlsm", moduleCode.Substring(0, 6), moduleCode);
                if (DocUtils.CheckExits(dmvtPath))
                {
                    DocUtils.DownloadFile("D:/ListDMVT", Path.GetFileName(dmvtPath), dmvtPath);
                    count++;
                }
            }
            MessageBox.Show(count.ToString());
        }

        void KiemTraVatTu()
        {
            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang kiem tra..."))
            {
                string[] listFilePath = Directory.GetFiles("D:/ListDMVT");
                foreach (string filePath in listFilePath)
                {
                    try
                    {
                        string moduleCode = Path.GetFileName(filePath).Substring(3, 10);
                        DataTable dtDMVT = TextUtils.ExcelToDatatableNoHeader(filePath, "DMVT");
                        var results = from myRow in dtDMVT.AsEnumerable()
                                      where TextUtils.ToDecimal(myRow.Field<string>("F1")) > 0
                                      && !(myRow.Field<string>("F4")).StartsWith("TPAD")
                                      && !(myRow.Field<string>("F4")).StartsWith("PCB")
                                      select myRow;
                        if (results.Count() == 0) continue;

                        dtDMVT = results.CopyToDataTable();

                        for (int i = 0; i <= dtDMVT.Rows.Count - 1; i++)
                        {
                            string nameDMVT = dtDMVT.Rows[i][3].ToString();
                            if (nameDMVT == "") continue;
                            string hang = dtDMVT.Rows[i][9].ToString();

                            List<string> errorString = new List<string>();

                            #region Kiem tra hang co hop le
                            DataTable dtGroup = TextUtils.Select("select CustomerCode from vMaterialCustomer a where replace(a.Code,' ','') ='" + nameDMVT.Replace(" ", "").Replace("(", "/") + "'");
                            if (dtGroup.Rows.Count > 0)
                            {
                                DataRow[] drsCustomer = dtGroup.Select("CustomerCode = '" + hang + "'");
                                if (drsCustomer.Count() == 0)
                                {
                                    errorString.Add("Hãng không được sử dụng");
                                }
                            }
                            #endregion

                            #region Vật tư ngừng sử dụng

                            DataTable dtViewMaterial = TextUtils.Select("SELECT MaterialGroupCode FROM [vMaterial] with (nolock) where replace(Code,' ','') = '" + nameDMVT.Replace(" ", "").Replace(")", "/") + "'");
                            if (dtViewMaterial.Rows.Count > 0)
                            {
                                string materialGroupCode = dtViewMaterial.Rows[0][0].ToString();
                                if (materialGroupCode == "TPAVT.Z01")
                                {
                                    errorString.Add("Vật tư ngừng sử dụng");
                                }
                            }
                            #endregion

                            #region Vật tư tạm dừng sử dụng
                            DataTable dtMaterialCSDL = TextUtils.Select("SELECT * FROM [Material] with (nolock) where [StopStatus] = 1 and replace(Code,' ','') = '" + nameDMVT.Replace(" ", "").Replace(")", "/") + "'");
                            if (dtMaterialCSDL.Rows.Count > 0)
                            {
                                errorString.Add("Vật tư tạm dừng sử dụng");
                            }
                            #endregion

                            #region Kiểm tra trên quản lý sản xuất
                            //Kiem tra xem vat tu co trong kho chua
                            DataTable dtMaterialQLSX = LibQLSX.Select("SELECT top 1 p.PartsCode,m.ManufacturerCode"
                                + " FROM Manufacturer m RIGHT OUTER JOIN"
                                + " PartsManufacturer pm ON m.ManufacturerId = pm.ManufacturerId RIGHT OUTER JOIN"
                                + " Parts p ON pm.PartsId = p.PartsId"
                                + " where p.PartsCode = '" + nameDMVT.Replace(" ", "").Replace("(", "/") + "'");
                            if (dtMaterialQLSX.Rows.Count == 0)
                            {
                                errorString.Add("Vật tư không tồn tại");
                            }
                            else
                            {
                                if (dtMaterialQLSX.Rows[0]["ManufacturerCode"].ToString().ToUpper() != hang.ToUpper())
                                {
                                    errorString.Add("Hãng khác với hãng trên QLSX (" + hang + " - " + dtMaterialQLSX.Rows[0]["ManufacturerCode"].ToString() + ")");
                                }
                            }
                            #endregion

                            if (errorString.Count > 0)
                            {
                                TestModel model = new TestModel();
                                model.ModuleCode = moduleCode;
                                model.MaterialCode = nameDMVT;
                                model.MaterialName = dtDMVT.Rows[i]["F2"].ToString();
                                model.Hang = hang;
                                model.Error = string.Join(", ", errorString.ToArray());
                                TestBO.Instance.Insert(model);
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            MessageBox.Show("ok");
        }

        void KiemTraVatTu_VatLieu()
        {
            DataTable dtError = new DataTable();
            dtError.Columns.Add("ModuleCode");
            dtError.Columns.Add("STT");
            dtError.Columns.Add("MaterialCode");
            dtError.Columns.Add("MaterialName");
            dtError.Columns.Add("MaVatLieu");
            dtError.Columns.Add("VatLieu");
            dtError.Columns.Add("Error");

            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang kiem tra..."))
            {
                string[] listFilePath = Directory.GetFiles("D:/ListDMVT");
                foreach (string filePath in listFilePath)
                {
                    try
                    {
                        string moduleCode = Path.GetFileName(filePath).Substring(3, 10);
                        DataTable dtDMVT = TextUtils.ExcelToDatatableNoHeader(filePath, "DMVT");

                        var results = from myRow in dtDMVT.AsEnumerable()
                                      where TextUtils.ToDecimal(myRow.Field<string>("F1") != "" && myRow.Field<string>("F1") != null 
                                      ? myRow.Field<string>("F1").Substring(0, 1) : "") > 0
                                      select myRow;

                        if (results == null) continue;
                        if (results.Count() == 0) continue;
                        if (results.Count() > 0)
                        {
                            dtDMVT = results.CopyToDataTable();
                        }                        

                        for (int i = 0; i <= dtDMVT.Rows.Count - 1; i++)
                        {
                            string nameDMVT = dtDMVT.Rows[i]["F4"].ToString();
                            if (nameDMVT == "") continue;
                            //string hang = dtDMVT.Rows[i][9].ToString();
                            string MaVatLieu = TextUtils.ToString(dtDMVT.Rows[i]["F5"]).Trim();
                            string VatLieu = TextUtils.ToString(dtDMVT.Rows[i]["F8"]).Trim();
                            string stt = TextUtils.ToString(dtDMVT.Rows[i]["F1"]).Trim();

                            List<string> errorString = new List<string>();

                            if (MaVatLieu != "")
                            {
                                DataTable dtMaVatLieu = LibQLSX.Select("SELECT top 1 PartsCode from Parts with(nolock)"
                               + " where PartsCode = N'" + MaVatLieu.Replace(" ", "").Replace("(", "/") + "'");
                                if (dtMaVatLieu.Rows.Count == 0)
                                {
                                    errorString.Add("Mã vật liệu không tồn tại trên QLSX");
                                }
                            }

                            if (VatLieu != "" && VatLieu != "-")
                            {
                                DataTable dtVatLieu = LibQLSX.Select("SELECT top 1 MaterialsId from MaterialsModel with(nolock)"
                              + " where MaterialsId = N'" + VatLieu + "'");
                                if (dtVatLieu.Rows.Count == 0)
                                {
                                    errorString.Add("Vật tư không tồn tại trên QLSX");
                                }
                            }

                            if (MaVatLieu != "" && (VatLieu == "" || VatLieu == "-"))
                            {
                                errorString.Add("Vật tư không có vật liệu");
                            }

                            if (errorString.Count > 0)
                            {                                
                                DataRow dr = dtError.NewRow();
                                dr["ModuleCode"] = moduleCode;
                                dr["STT"] = stt;
                                dr["MaterialCode"] = nameDMVT;
                                dr["MaterialName"] = TextUtils.ToString(dtDMVT.Rows[i]["F2"]);
                                dr["MaVatLieu"] = MaVatLieu;
                                dr["VatLieu"] = VatLieu;
                                dr["Error"] = string.Join(", ", errorString.ToArray());
                                dtError.Rows.Add(dr);
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            gridControl1.DataSource = dtError;
            MessageBox.Show("ok");
        }

        void runMacroInWord(string filePath, string macroName, string parameter)
        {
            Word.Application word = new Word.Application();
            Word.Document doc = new Word.Document();

            doc = word.Documents.Open(filePath);
            word.Run(macroName, parameter);
            doc.Save();
            doc.Close();
            word.Quit();

            Process.Start(filePath);
        }

        void loadAllDmvt()
        {
            string[] listFile = Directory.GetFiles("D:\\ListDMVT", "*.*", SearchOption.TopDirectoryOnly);
            DataTable dtAll = new DataTable();
            foreach (string filePath in listFile)
            {
                try
                {
                    DataTable dt = TextUtils.ExcelToDatatableNoHeader(filePath, "DMVT");
                    dt = dt.AsEnumerable()
                                .Where(row => TextUtils.ToInt(row.Field<string>("F1") == "" ||
                                    row.Field<string>("F1") == null ? "0" : row.Field<string>("F1").Substring(0, 1)) > 0)
                                .CopyToDataTable();
                    if (dtAll.Rows.Count == 0)
                    {
                        dtAll = dt.Copy();
                    }
                    else
                    {
                        dtAll.Merge(dt);
                    }
                }
                catch
                {
                }
            }
            gridControl1.DataSource = dtAll;
            gridView1.BestFitColumns();
        }

        void addMaterialModuleLink()
        {
            string localPath = "D:\\ListDMVT";
            string[] listFiles = Directory.GetFiles(localPath);
            foreach (string filePath in listFiles)
            {
                string fileName = Path.GetFileName(filePath);
                string moduleCode = fileName.Substring(3, 10);
                try
                {
                    DataTable dtDMVT = TextUtils.ExcelToDatatableNoHeader(filePath, "DMVT");
                    string designer = TextUtils.ToString(dtDMVT.Rows[3]["F3"]);

                    DataRow[] drs = dtDMVT.AsEnumerable()
                                .Where(row => TextUtils.ToInt(row.Field<string>("F1") == "" ||
                                    row.Field<string>("F1") == null ? "0" : row.Field<string>("F1").Substring(0, 1)) > 0)
                                .ToArray();

                    foreach (DataRow row in drs)
                    {
                        MaterialModuleLinkModel link = new MaterialModuleLinkModel();
                        link.ModuleCode = moduleCode;
                        link.STT = TextUtils.ToString(row["F1"]);
                        link.Name = TextUtils.ToString(row["F2"]);
                        link.ThongSo = TextUtils.ToString(row["F3"]);
                        link.Code = TextUtils.ToString(row["F4"]);
                        link.MaVatLieu = TextUtils.ToString(row["F5"]);
                        link.Unit = TextUtils.ToString(row["F6"]);
                        link.Qty = TextUtils.ToDecimal(row["F7"]);
                        link.VatLieu = TextUtils.ToString(row["F8"]);
                        link.Weight = TextUtils.ToDecimal(row["F9"]);
                        link.Hang = TextUtils.ToString(row["F10"]);
                        link.Note = TextUtils.ToString(row["F11"]);
                        link.Designer = designer;
                        link.DateCreated = DateTime.Now;

                        MaterialModuleLinkBO.Instance.Insert(link);
                    }
                }
                catch
                {                  
                }                
            }
        }

        void addSuppliersManufacturerLink()
        {
            DataTable dt = LibQLSX.Select("SELECT     vRequirePartFull.SupplierId, Manufacturer.ManufacturerId"
                                            +" FROM vRequirePartFull INNER JOIN"
                                            +" Manufacturer ON vRequirePartFull.ManufacturerCode = Manufacturer.ManufacturerCode" 
                                            +" order by vRequirePartFull.SupplierId");
            long count = 0;
            foreach (DataRow r in dt.Rows)
            {
                string supplierId = TextUtils.ToString(r["SupplierId"]);
                string manufacturerId = TextUtils.ToString(r["ManufacturerId"]);
                if (supplierId == "" || manufacturerId == "") continue;

                DataTable dtSuppliersManufacturerLink = LibQLSX.Select("select * from SuppliersManufacturerLink with(nolock) where SupplierId = '" 
                    + supplierId + "' and ManufacturerId = '" + manufacturerId + "'");
                if (dtSuppliersManufacturerLink.Rows.Count > 0) continue;
                SuppliersManufacturerLinkModel model = new SuppliersManufacturerLinkModel();
                model.SupplierId = supplierId;
                model.ManufacturerId = manufacturerId;
                SuppliersManufacturerLinkBO.Instance.Insert(model);
                count++;
            }
            MessageBox.Show("Insert: " + count);
        }

        void uploadFolder(string folderPath, string ftpPath)
        {
            DocUtils.InitFTPTK();
            string folderName = Path.GetFileName(folderPath);
            if (!DocUtils.CheckExits(ftpPath + "/" + folderName))
            {
                DocUtils.MakeDir(ftpPath + "/" + folderName);
            }

            DocUtils.UploadDirectory(folderPath, ftpPath + "/" + folderName);
        }

        void updatePart()
        {
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                return;
            }
            //filePath = @"E:\PROJECTS\TanPhat\KAIZEN\Phòng vật tư\HScode-Xuất nhập khẩu\MISUMI-HSCODE.xlsx";
            DataTable dt = TextUtils.ExcelToDatatableNoHeader(filePath, "Sheet2");
            dt = dt.Select("F2 is not null and F2 <> ''").CopyToDataTable();
            DataTable dtPart = LibQLSX.Select("select * from Parts with(nolock)");

            int count = 0;
            foreach (DataRow row in dt.Rows)
            {
                string partsCode = TextUtils.ToString(row["F2"]);
                if (partsCode == "") continue;
                string des = TextUtils.ToString(row["F3"]);
                string hsCode = TextUtils.ToString(row["F4"]);
                decimal importTax = TextUtils.ToDecimal(row["F5"]);

                DataRow[] drs = dtPart.Select("PartsCode = '" + partsCode.Trim() + "'");
                if (drs.Length > 0)
                {
                    string partsId = TextUtils.ToString(drs[0]["PartsId"]);

                    PartsModel part = (PartsModel)PartsBO.Instance.FindByAttribute("PartsId", partsId)[0];
                    part.Description = des;
                    part.HsCode = hsCode;
                    part.ImportTax = importTax;

                    PartsBO.Instance.UpdateQLSX(part);
                    count++;
                }
            }
            MessageBox.Show(count.ToString());
        }

        void upateDatePrice()
        {
            string filePath = @"E:\PROJECTS\TanPhat\KAIZEN\VAT TU DA THEM GIA.xlsx";
            DataTable dt = TextUtils.ExcelToDatatableNoHeader(filePath, "Sheet");
            dt = dt.Select("F4 is not null and F4 <> ''").CopyToDataTable();
            DataTable dtPart = LibQLSX.Select("select * from Parts with(nolock)");

            int count = 0;
            foreach (DataRow row in dt.Rows)
            {
                string partsCode = TextUtils.ToString(row["F4"]);
                string MaVatLieu = TextUtils.ToString(row["F5"]);
                if (partsCode == "") continue;
                if (MaVatLieu != "" && MaVatLieu != "GCCX")
                {
                    partsCode = MaVatLieu;
                }

                DataRow[] drs = dtPart.Select("PartsCode = '" + partsCode.Trim() + "'");
                if (drs.Length > 0)
                {
                    string partsId = TextUtils.ToString(drs[0]["PartsId"]);
                    decimal price = TextUtils.ToDecimal(drs[0]["Price"]);
                    if (price <= 1) continue;

                    PartsModel part = (PartsModel)PartsBO.Instance.FindByAttribute("PartsId", partsId)[0];
                    part.UpdatedPriceDate = DateTime.Now;
                    PartsBO.Instance.UpdateQLSX(part);
                    count++;
                }
            }
            MessageBox.Show(count.ToString());
        }

        void updateUser()
        {
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                return;
            }
            DataTable dt = TextUtils.ExcelToDatatableNoHeader(filePath, "Sheet2");
            dt = dt.Select("F2 is not null and F2 <> ''").CopyToDataTable();
            DataTable dtUser = LibQLSX.Select("select * from Users with(nolock)");

            int count = 0;
            foreach (DataRow row in dt.Rows)
            {
                try
                {
                    string code = TextUtils.ToString(row["F2"]);
                    if (code == "") continue;
                    string userName = TextUtils.ToString(row["F3"]);

                    DataRow[] drs = dtUser.Select("UserName = '" + userName.Trim() + "'");
                    if (drs.Length > 0)
                    {
                        string userId = TextUtils.ToString(drs[0]["UserId"]);

                        TPA.Model.UsersModel user = (TPA.Model.UsersModel)TPA.Business.UsersBO.Instance.FindByAttribute("UserId", userId)[0];
                        user.Code = code;

                        PartsBO.Instance.UpdateQLSX(user);
                        count++;
                    }
                }
                catch 
                {
                }
            }
            MessageBox.Show(count.ToString());
        }


        private string ExtractTextFromImage()
        {
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                return "";
            }
            Document modiDocument = new Document();
            modiDocument.Create(filePath);
            modiDocument.OCR(MiLANGUAGES.miLANG_ENGLISH);
            MODI.Image modiImage = (modiDocument.Images[0] as MODI.Image);
            string extractedText = modiImage.Layout.Text;
            modiDocument.Close();
            return extractedText;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            textBox1.Text = ExtractTextFromImage();
            //updatePart();
            //upateDatePrice();
            //updateUser();
        }

        private void btnDownloadDMVT_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn download tất cả các dmvt ở trên nguồn?", TextUtils.Caption, MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                DownloadDMVT();
            }
        }

        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            //TextUtils.ExportExcel(gridView1);
        }

        private void btnBaoGia_Click(object sender, EventArgs e)
        {
            frmBaoGia frm = new frmBaoGia();
            TextUtils.OpenForm(frm);
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (gridView2.RowCount > 0)
            {
                TextUtils.ExportExcel(gridView2);
            }
        }

        private void btnGetBarcode_Click(object sender, EventArgs e)
        {
            string filePath = @"D:\Thietke.Ck\TPAD.M\TPAD.M5211.Ck\BCCk.TPAD.M5211\BC-CAD.TPAD.M5211\TPAD.M5211.01.01.01.jpg";
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog()==DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            System.Drawing.Image img = System.Drawing.Image.FromFile(filePath);
            Bitmap mBitmap = new Bitmap(img);

            ArrayList barcodes = new ArrayList();
            BarcodeImaging.FullScanPage(ref barcodes, mBitmap, 100);
            mBitmap.Dispose();
            if (barcodes.Count>0)
            {
                textBox1.Text = barcodes.ToArray()[0].ToString();
            }
            else
            {
                MessageBox.Show("Không tìm thấy barcode!");
            }
           
        }

        private void btnGetTextInImage_Click(object sender, EventArgs e)
        {
            //string filePath = "";
            //OpenFileDialog ofd = new OpenFileDialog();
            //if (ofd.ShowDialog() == DialogResult.OK)
            //{
            //    textBox1.Text = TextUtils.ExtractTextFromPdf(ofd.FileName);
            //}
        }

        private void btnLoadProductModule_Click(object sender, EventArgs e)
        {
            loadTree("T009.1501");
        }

        private void treeData_CellValueChanged(object sender, DevExpress.XtraTreeList.CellValueChangedEventArgs e)
        {

        }

        private void btnGetListNoPrice_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            TextUtils.LoadModulePriceTPAD("TPAD.B3580", false, out dt);
            gridControl1.DataSource = dt;
        }

        private void btnAddMaterialModuleLink_Click(object sender, EventArgs e)
        {
            //addMaterialModuleLink();
        }
    }
}
