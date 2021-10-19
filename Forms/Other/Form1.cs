using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.VisualBasic;
using BMS.Model;
using BMS.Business;
using DevExpress.Utils;
using Forms.Properties;
using System.Collections;
using BMS.Utils;
//using Barcode;
using System.Drawing.Imaging;
using System.Net.Mail;
using System.Net;

namespace BMS
{
    public partial class Form1 : _Forms
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable GetDMVT(string filePath)
        {
            try
            {
                DataTable dt1 = TextUtils.ExcelToDatatable(filePath.Trim(), "DMVT");
                DataTable dt = dt1.Copy();
                for (int i = 0; i <= 4; i++)
                {
                    dt.Rows.RemoveAt(0);
                }
                //for (int i = 0; i <4; i++)
                //{
                //    dt.Rows.RemoveAt(dt.Rows.Count - 1);
                //}
                for (int i = 1; i <= 5; i++)
                {
                    dt.Columns.RemoveAt(4);
                }
                dt.Columns.RemoveAt(dt.Columns.Count - 1);
                dt.Columns.RemoveAt(dt.Columns.Count - 1);
                return dt1;
            }
            catch (Exception)
            {
                return new DataTable();
            }
        }

        public DataTable GetDMVT(DataTable dt)
        {
            try
            {                
                DataTable dt1 = dt.Copy();
                for (int i = 0; i <= 4; i++)
                {
                    dt1.Rows.RemoveAt(0);
                }
                return dt1.Select("F1<>''").CopyToDataTable();
            }
            catch (Exception)
            {
                return new DataTable();
            }
        }

        void UpdateFTP()
        {
            string initPath = @"F:\data\SQL SXLR\DataTPA\Sourcecode\";
            string cokhiPath = initPath + "/Thietke.Ck/";
            string dienTuPath = initPath + "/Thietke.Dt/PCB/";
			string dienPath = initPath + "/Thietke.Dn/";
            ArrayList array = ModulesBO.Instance.FindByExpression(new Expression("Status", 2).And(new Expression("Code", "TPAD", "like")));
            foreach (var row in array)
            {
                ModulesModel module = (ModulesModel)row;
                string pathDMVT = cokhiPath + "//" + module.Code.Substring(0, 6) + "//" + module.Code + ".Ck//VT." + module.Code + ".xlsm";
                if (!File.Exists(pathDMVT)) continue;
                FileInfo fi = new FileInfo(pathDMVT);                

                DataTable dtDMVT = TextUtils.ExcelToDatatable(pathDMVT, "DMVT");
                DataTable dtListMaterial = GetDMVT(dtDMVT);

                DataTable dtLink = TextUtils.Select("MaterialModuleLink", new Expression("ModuleID", module.ID));

                string filePathIam = cokhiPath + module.Code.Substring(0, 6) + "/" + module.Code + ".Ck/3D." + module.Code + "/" + module.Code + ".iam";
                string folderPahModuleImage = "";//Hỏi anh Hậu????????????????????????????????

                #region UPDATE TÀI LIỆU THIẾT KẾ CHO SẢN PHẨM
                try
                {
                    int DT = 0;
                    int D = 0;

                    //string dienPath = "D:\\Thietke.Dn\\" + module.Code.Substring(0, 6) + "\\" + module.Code + ".Dn";
                    if (Directory.Exists(dienPath))
                    {
                        D = 1;
                    }

                    try
                    {
                        if (dtListMaterial.Select("F4 like '%PCB%'").Count() > 0)
                        {
                            DT = 1;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    module.FileElectric = D;
                    module.FileElectronic = DT;
                    module.FileMechanics = 1;
                    ModulesBO.Instance.Update(module);
                }
                catch (Exception)
                {
                }
                #endregion UPDATE TÀI LIỆU THIẾT KẾ CHO SẢN PHẨM

                if (dtLink.Rows.Count > 0)
                {
                    int size = TextUtils.ToInt(dtLink.Rows[0]["Size"]);
                    if (size != (int)(fi.Length))
                    {
                        #region Tạo ảnh của sản phẩm
                        try
                        {
                            //IPTDetail.LoadData(filePathIam);
                            //Bitmap bit = new Bitmap(IPTDetail.Image, IPTDetail.Image.Width / 3, IPTDetail.Image.Height / 3);
                            //bit.Save(folderPahModuleImage + "/" + module.Code+".png", ImageFormat.Png);

                            //module.ImagePath = "\\SERVER\\data2\\Thietke\\PHANMEM\\ModuleImage\\" + module.Code + ".png";
                            //ModulesBO.Instance.Update(module);
                        }
                        catch
                        {                           
                        }
                        #endregion

                        #region THÊM NGƯỜI THIẾT KẾ
                        try
                        {
                            ModuleDesignerBO.Instance.DeleteByAttribute("ModuleID", module.ID);
                            if (dtDMVT.Rows[2]["F3"].ToString()!="")
                            {
                                ModuleDesignerModel model = new ModuleDesignerModel();
                                model.LoginName = dtDMVT.Rows[2]["F3"].ToString();
                                model.ModuleID = module.ID;
                                model.WorkDetail = "Thiết kế cơ khí: " + fi.LastWriteTime.ToString("dd/MM/yyyy");
                                model.GroupType = 0;
                                model.Author = dtDMVT.Rows[2]["F3"].ToString();
                                ModuleDesignerBO.Instance.Insert(model);
                            }                            

                            try
                            {
                                if (dtListMaterial.Select("F4 like '%PCB%'").Count() > 0)
                                {
                                    foreach (DataRow item in dtListMaterial.Select("F4 like '%PCB%'"))
                                    {
                                        string code = item["F4"].ToString();
                                        if (code!="")
                                        {
                                            string pathDMVT_DT = dienTuPath + "/" + code + "/PRD." + code + "/VT." + code + ".xls";
                                            DataTable dtDMVT_DT = TextUtils.ExcelToDatatable(pathDMVT_DT, "DMVT");

                                            ModuleDesignerModel modelDT = new ModuleDesignerModel();
                                            modelDT.LoginName = dtDMVT_DT.Rows[2]["F3"].ToString();
                                            modelDT.ModuleID = module.ID;
                                            modelDT.WorkDetail = "Thiết kế điện tử: " + code;
                                            modelDT.GroupType = 2;
                                            modelDT.Author = dtDMVT_DT.Rows[2]["F3"].ToString();
                                            ModuleDesignerBO.Instance.Insert(modelDT);
                                        }                                        
                                    }
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
                        #endregion THÊM NGƯỜI THIẾT KẾ

                        #region TẠO DANH MỤC VẬT TƯ CHO MODULE
                        //xóa danh sách vật tư cũ
                        MaterialModuleLinkBO.Instance.DeleteByAttribute("ModuleID", module.ID);
                        try
                        {
                            foreach (DataRow item in dtListMaterial.Rows)
                            {
                                MaterialModuleLinkModel linkModel = new MaterialModuleLinkModel();

                                linkModel.ModuleID = module.ID;

                                //linkModel.DateCreated = fi.LastWriteTime.ToString("dd/MM/yyyy");
                                linkModel.Designer = dtDMVT.Rows[2]["F3"].ToString();
                                linkModel.SizeTK = (int)(fi.Length);
                                linkModel.Author = dtDMVT.Rows[2]["F3"].ToString();

                                linkModel.STT = item["F1"].ToString();
                                linkModel.Name = item["F2"].ToString();
                                linkModel.ThongSo = item["F3"].ToString();
                                linkModel.Code = item["F4"].ToString();
                                linkModel.MaVatLieu = item["F5"].ToString();
                                linkModel.Unit = item["F6"].ToString();
                                linkModel.Qty = TextUtils.ToInt(item["F7"].ToString());
                                linkModel.VatLieu = item["F8"].ToString();
                                linkModel.Weight = TextUtils.ToDecimal(item["F9"].ToString());
                                linkModel.Hang = item["F10"].ToString();
                                try
                                {
                                    ArrayList listMaterial = MaterialBO.Instance.FindByAttribute("Code", linkModel.Code);
                                    MaterialModel model = (MaterialModel)listMaterial[0];
                                    linkModel.Price = model.Price;
                                    linkModel.Delivery = model.DeliveryPeriod;
                                }
                                catch (Exception)
                                {
                                    linkModel.Price = 0;
                                    linkModel.Delivery = 0;
                                }

                                MaterialModuleLinkBO.Instance.Insert(linkModel);
                            }
                        }
                        catch (Exception)
                        {
                        }
                        #endregion TẠO DANH MỤC VẬT TƯ CHO MODULE
                    }
                }
                else
                {
                    #region Tạo ảnh của sản phẩm
                    try
                    {
                        //IPTDetail.LoadData(filePathIam);
                        //Bitmap bit = new Bitmap(IPTDetail.Image, IPTDetail.Image.Width / 3, IPTDetail.Image.Height / 3);
                        //bit.Save(folderPahModuleImage + "/" + module.Code + ".png", ImageFormat.Png);

                        //module.ImagePath = "\\SERVER\\data2\\Thietke\\PHANMEM\\ModuleImage\\" + module.Code + ".png";
                        //ModulesBO.Instance.Update(module);
                    }
                    catch
                    {
                    }
                    #endregion

                    #region THÊM NGƯỜI THIẾT KẾ
                    try
                    {
                        ModuleDesignerModel model = new ModuleDesignerModel();
                        model.LoginName = dtDMVT.Rows[2]["F3"].ToString();
                        model.ModuleID = module.ID;
                        model.WorkDetail = "Thiết kế cơ khí: " + fi.LastWriteTime.ToString("dd/MM/yyyy");
                        model.GroupType = 0;
                        model.Author = dtDMVT.Rows[2]["F3"].ToString();
                        ModuleDesignerBO.Instance.Insert(model);

                        try
                        {
                            if (dtListMaterial.Select("F4 like '%PCB%'").Count() > 0)
                            {
                                foreach (DataRow item in dtListMaterial.Select("F4 like '%PCB%'"))
                                {
                                    string code = item["F4"].ToString();
                                    if (code != "")
                                    {
                                        string pathDMVT_DT = dienTuPath + "/" + code + "/PRD." + code + "/VT." + code + ".xls";
                                        DataTable dtDMVT_DT = TextUtils.ExcelToDatatable(pathDMVT_DT, "DMVT");

                                        ModuleDesignerModel modelDT = new ModuleDesignerModel();
                                        modelDT.LoginName = dtDMVT_DT.Rows[2]["F3"].ToString();
                                        modelDT.ModuleID = module.ID;
                                        modelDT.WorkDetail = "Thiết kế điện tử: " + code;
                                        modelDT.GroupType = 2;
                                        modelDT.Author = dtDMVT_DT.Rows[2]["F3"].ToString();
                                        ModuleDesignerBO.Instance.Insert(modelDT);
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    catch (Exception)
                    {
                    }
                    #endregion THÊM NGƯỜI THIẾT KẾ

                    #region TẠO DANH MỤC VẬT TƯ CHO MODULE
                    try
                    {
                        foreach (DataRow item in dtListMaterial.Rows)
                        {
                            MaterialModuleLinkModel linkModel = new MaterialModuleLinkModel();

                            linkModel.ModuleID = module.ID;

                            //linkModel.DateCreated = fi.LastWriteTime.ToString("dd/MM/yyyy");
                            linkModel.Designer = dtDMVT.Rows[2]["F3"].ToString();
                            linkModel.SizeTK = (int)(fi.Length);
                            linkModel.Author = dtDMVT.Rows[2]["F3"].ToString();

                            linkModel.STT = item["F1"].ToString();
                            linkModel.Name = item["F2"].ToString();
                            linkModel.ThongSo = item["F3"].ToString();
                            linkModel.Code = item["F4"].ToString();
                            linkModel.MaVatLieu = item["F5"].ToString();
                            linkModel.Unit = item["F6"].ToString();
                            linkModel.Qty = TextUtils.ToInt(item["F7"].ToString());
                            linkModel.VatLieu = item["F8"].ToString();
                            linkModel.Weight = TextUtils.ToDecimal(item["F9"].ToString());
                            linkModel.Hang = item["F10"].ToString();

                            try
                            {
                                ArrayList listMaterial = MaterialBO.Instance.FindByAttribute("Code", linkModel.Code);
                                MaterialModel model = (MaterialModel)listMaterial[0];
                                linkModel.Price = model.Price;
                                linkModel.Delivery = model.DeliveryPeriod;
                            }
                            catch (Exception)
                            {
                                linkModel.Price = 0;
                                linkModel.Delivery = 0;
                            }

                            MaterialModuleLinkBO.Instance.Insert(linkModel);
                        }
                    }
                    catch (Exception)
                    {
                    }
                    #endregion TẠO DANH MỤC VẬT TƯ CHO MODULE
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        void CreateExcel()
        {           
            string filePath = "D:\\tet.xlsx";

            Excel.Application app = default(Excel.Application);
            Excel.Workbook workBoook = default(Excel.Workbook);
            Excel.Worksheet workSheet = default(Excel.Worksheet);
            try
            {
                app = new Excel.Application();
                app.Workbooks.Open(filePath);
                workBoook = app.Workbooks[1];
                workSheet = (Excel.Worksheet)workBoook.Worksheets[1];

                //Excel.Range oRng = workSheet.Range["M1"];
                //oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                for (int i = 0; i < 3; i++)
                {
                    ((Excel.Range)workSheet.Rows[13]).Insert();
                }

                Excel.Range rCopy = workSheet.Range["A10","Z12"];
                rCopy.Copy();

                Excel.Range rPaste = workSheet.Range["A14", "Z16"];
                rPaste.PasteSpecial(Excel.XlPasteType.xlPasteAll);

                DataTable dt=TextUtils.Select("SELECT top 1 (select COUNT(*) from [dbo].[fnStringToTable]([NameTS],',')) as countTS "
	                                        +" FROM [QLKHCV].[dbo].[MaterialHistory] where ModuleID = "
	                                        +" order by (select COUNT(*) from [dbo].[fnStringToTable]([NameTS],',')) desc)");
                app.ActiveWorkbook.Save();
                app.Workbooks.Close();
                app.Quit();

                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                TextUtils.ReleaseComObject(app);
                TextUtils.ReleaseComObject(workBoook);
                TextUtils.ReleaseComObject(workSheet);
            }
            
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            //CreateExcel();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang download file CAD tổng..."))
                {
                    foreach (string filePathIAM in ofd.FileNames)
                    {
                        try
                        {
                            IPTDetail.LoadData(filePathIAM);
                            Image img = IPTDetail.Image;
                            //img.Save(@"\\SERVER\data2\Dulieu_TPA\PHANMEM\ModuleImage\" + Module.Code + ".jpg", ImageFormat.Jpeg);
                            //@"\\SERVER\data2\Dulieu_TPA\PHANMEM\ModuleImage\" + Module.Code + ".jpg";
                        }
                        catch (Exception)
                        {                           
                        }
                    }
                }
                //IPTDetail.LoadData(ofd.FileName);
                //Image img = IPTDetail.Image;
                ////img.Save("\\SERVER\\data2\\Thietke\\PHANMEM\\ModuleImage\\" + Path.GetFileNameWithoutExtension(ofd.FileName) + ".jpg", ImageFormat.Jpeg);
                //img.Save("\\SERVER\\data2\\Dulieu_TPA\\PHANMEM\\ModuleImage\\"+ Path.GetFileNameWithoutExtension(ofd.FileName) + ".jpg", ImageFormat.Jpeg);
                ////pictureBox1.Image = IPTDetail.Image;

            }
            MessageBox.Show("Thành công!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TextUtils.ExcuteProcedure("spDongBoMaterialQLSX", null, null);
            MessageBox.Show("OK");
        }

        //upload các thông tin của model trên máy chủ
        private void button3_Click(object sender, EventArgs e)
        {
            UpdateFTP();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmCopyExcel frm = new frmCopyExcel();
            TextUtils.OpenForm(frm);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            List<int> listPermission = new List<int>() { 23, 43, 76 };
            DataTable dt = TextUtils.Select("select * from vUserInfo where DepartMentId<>1 and id <> 17");
            foreach (DataRow row in dt.Rows)
            {
                int userID = TextUtils.ToInt(row[0]);
                foreach (int item in listPermission)
                {
                    UserRightDistributionModel model = new UserRightDistributionModel();
                    model.UserID = userID;
                    model.FormAndFunctionID = item;
                    UserRightDistributionBO.Instance.Insert(model);
                }
            }
            MessageBox.Show("OK");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string initPath = @"D:\ErrorTree\";
            DataTable dtError = TextUtils.Select("Select * from ModuleError");
            foreach (DataRow row in dtError.Rows)
            {
                int moduleID = TextUtils.ToInt(row["ModuleID"]);
                string moduleCode = ((ModulesModel)ModulesBO.Instance.FindByPK(moduleID)).Code;
                string path = initPath + moduleCode;
                Directory.CreateDirectory(path);

                Directory.CreateDirectory(path + "/" + row["Code"].ToString());
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog pfd = new OpenFileDialog();
            pfd.Multiselect = false;
            if (pfd.ShowDialog() == DialogResult.OK)
            {
                //DataTable _dtHistoryCheck = TextUtils.Select("CheckCTTKHistory", new Expression("ID", 0));
                string filePath = pfd.FileName;

                //System.Collections.ArrayList barcodes = new System.Collections.ArrayList();
                Image img = Image.FromFile(filePath);
                Bitmap mBitmap = new Bitmap(img);

                //string[] a = BarcodeReader.read(mBitmap, BarcodeReader.CODE39);

                //string[] b = BarcodeReader.read(mBitmap, BarcodeReader.CODE39);
                textBox1.Text = "";
                //string[] c = OnBarcode.Barcode.BarcodeScanner.BarcodeScanner.Scan(mBitmap, OnBarcode.Barcode.BarcodeScanner.BarcodeType.Code39);
                //foreach (string item in c)
                //{
                //    textBox1.Text += item + Environment.NewLine;
                //}
            }
        }

        /// <summary>
        /// Download Bản Cad tổng của tất cả các thiết kế
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            DocUtils.InitFTPQLSX();
            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang download file CAD tổng..."))
            {
                ArrayList listModules = ModulesBO.Instance.FindByExpression(new Expression("Status", 2).And(new Expression("Code","TPAD","like")));
                string pathLocal = "D:/File Tong IAM";
                Directory.CreateDirectory(pathLocal);
                foreach (var item in listModules)
                {
                    ModulesModel module = (ModulesModel)item;
                    string code = module.Code;
                    string filePathCAD = string.Format(@"Thietke.Ck\{0}\{1}.Ck\3D.{1}\{1}.iam", code.Substring(0, 6), code);

                    try
                    {
                        DocUtils.DownloadFile(pathLocal, code + ".iam", filePathCAD);
                    }
                    catch
                    {                        
                    }                    
                }
            }
            MessageBox.Show("Thành công!");
        }


        /// <summary>
        /// Thêm ảnh lỗi cho các thiết kế
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            DocUtils.InitFTPTK();

            string path = "D:/ErrorTree";
            string[] listModulePath = Directory.GetDirectories(path);
            foreach (string modulePath in listModulePath)
            {
                string moduleCode = Path.GetFileName(modulePath);
                ModulesModel thisModule = (ModulesModel)ModulesBO.Instance.FindByExpression(new Expression("Code", moduleCode))[0];

                string[] listErrorPath = Directory.GetDirectories(modulePath);
                foreach (string errorPath in listErrorPath)
                {
                    string errorCode = Path.GetFileName(errorPath);                    
                    DataTable dtError = TextUtils.Select("ModuleError", new Expression("Code", errorCode));
                    ModuleErrorModel thisError;
                    if (dtError.Rows.Count > 0)
                    {
                        thisError = (ModuleErrorModel)ModuleErrorBO.Instance.FindByPK(TextUtils.ToInt(dtError.Rows[0]["ID"]));
                    }
                    else
                    {
                        thisError = new ModuleErrorModel();
                        thisError.ModuleID = thisModule.ID;
                        thisError.Code = errorCode;
                        thisError.Status = 0;
                        thisError.Description = "Chưa khai báo";
                        thisError.Author = Global.AppUserName;
                        thisError.UserID = Global.UserID;
                        ModuleErrorBO.Instance.Insert(thisError);
                    }

                    foreach (string errorFilePath in Directory.GetFiles(errorPath))
                    {
                        ProcessTransaction pt = new ProcessTransaction();
                        pt.OpenConnection();
                        pt.BeginTransaction();
                        try
                        {
                            FileInfo fInfo = new FileInfo(errorFilePath);

                            ModuleErrorImageModel errorImageModel = new ModuleErrorImageModel();
                            errorImageModel.DateCreated = TextUtils.GetSystemDate();
                            errorImageModel.FileName = Path.GetFileName(errorFilePath);
                            errorImageModel.Size = fInfo.Length;
                            errorImageModel.FilePath = "Modules\\ErrorImage\\" + errorImageModel.FileName;
                            errorImageModel.ModuleErrorID = thisError.ID;
                            errorImageModel.ID = (int)pt.Insert(errorImageModel);

                            DocUtils.UploadFile(errorFilePath, "Modules\\ErrorImage");

                            pt.CommitTransaction();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            pt.CloseConnection();
                        }
                    }
                }
            }
            MessageBox.Show("OK");
        }

        /// <summary>
        /// Phân bổ dữ liệu cho phòng kế toán
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            DocUtils.InitFTPQLSX();
            string _productCode = "TPAD.A2106";
            string _projectCode = "DA1";
            string _serverPathCK = "Thietke.Ck/" + _productCode.Substring(0, 6) + "/" + _productCode + ".Ck/";
            DataTable dt = TextUtils.Select("vPBDL_Structure_File", new Expression("DepartmentID", 8));
            ArrayList listStruc = PBDLStructureBO.Instance.FindByAttribute("DepartmentID", 8);
            if (listStruc.Count > 0)
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    string initPath = fbd.SelectedPath;

                    #region Lấy dữ liệu trong file danh mục vật tư cơ khí
                    try
                    {
                        if (DocUtils.CheckExits(_serverPathCK + "/VT." + _productCode + ".xlsm"))
                        {
                            //Download file danh mục vật tư
                            DocUtils.DownloadFile(initPath, "VT." + _productCode + ".xlsm", _serverPathCK + "/VT." + _productCode + ".xlsm");
                        }
                    }
                    catch
                    {
                    }

                    string pathDMVT = initPath + "/VT." + _productCode + ".xlsm";
                    DataTable dtDMVT = TextUtils.ExcelToDatatable(pathDMVT, "DMVT");

                    for (int i = 0; i <= 4; i++)
                    {
                        dtDMVT.Rows.RemoveAt(0);
                    }
                    File.Delete(pathDMVT);

                    #endregion

                    foreach (var item in listStruc)
                    {
                        PBDLStructureModel strucModel = (PBDLStructureModel)item;
                        string pathStruc = strucModel.Path.Replace("project", _projectCode).Replace("code", _productCode);
                        string pathLocal = initPath + "/" + pathStruc;
                        Directory.CreateDirectory(pathLocal);
                        DataTable dtFile = TextUtils.Select("vPBDL_Structure_File", new Expression("PBDLStructureID", strucModel.ID));
                        foreach (DataRow row in dtFile.Rows)
                        {
                            int pBDLFileID = TextUtils.ToInt(row["PBDLFileID"]);
                            PBDLFileModel fileModel = (PBDLFileModel)PBDLFileBO.Instance.FindByPK(pBDLFileID);
                            
                            if (fileModel.GetType == 0)//nhieu file
                            {
                                string pathServer = fileModel.FolderContain.Replace("group", _productCode.Substring(0, 6)).Replace("code", _productCode);

                                #region Mã vật liệu
                                if (fileModel.FilterMaVatLieu)//nếu có mã vật liệu
                                {
                                    DataRow[] listMaVatLieu = dtDMVT.Select("F5 <> '' and F4 like '%TPAD%' ");
                                    foreach (DataRow rowMVL in listMaVatLieu)
                                    {
                                        string fileName = rowMVL["F4"].ToString();
                                        try
                                        {
                                            if (DocUtils.CheckExits(pathServer + "/" + fileName + ".dwg"))
                                            {
                                                DocUtils.DownloadFile(pathLocal, fileName + fileModel.Extension, pathServer + "/" + fileName + fileModel.Extension);
                                            }
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                #endregion

                                #region Đơn vị
                                if (fileModel.FilterDonVi != "")
                                {
                                    DataRow[] listDonVi = dtDMVT.Select("F6 like '%" + fileModel.FilterDonVi.ToUpper() + "%' and F4 like '%TPAD%' ");
                                    foreach (DataRow rowMVL in listDonVi)
                                    {
                                        string fileName = rowMVL["F4"].ToString();
                                        try
                                        {
                                            if (DocUtils.CheckExits(pathServer + "/" + fileName + ".dwg"))
                                            {
                                                DocUtils.DownloadFile(pathLocal, fileName + fileModel.Extension, pathServer + "/" + fileName + fileModel.Extension);
                                            }
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                #endregion Đơn vị

                                #region Thông Số
                                if (fileModel.FilterThongSo != "")
                                {
                                    DataRow[] listThongSo = dtDMVT.Select("F3 like '%" + fileModel.FilterThongSo.ToUpper() + "%' and F4 like '%TPAD%' ");
                                    foreach (DataRow rowThongSo in listThongSo)
                                    {
                                        string fileName = rowThongSo["F4"].ToString();
                                        try
                                        {
                                            if (DocUtils.CheckExits(pathServer + "/" + fileName + ".dwg"))
                                            {
                                                DocUtils.DownloadFile(pathLocal, fileName + fileModel.Extension, pathServer + "/" + fileName + fileModel.Extension);
                                            }
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                #endregion Thông Số

                                #region MẶT
                                if (fileModel.MAT)
                                {
                                    try
                                    {
                                        string[] listFileMat = DocUtils.GetFilesList(pathServer);
                                        if (listFileMat!=null)
                                        {
                                            foreach (string fileName in listFileMat)
                                            {
                                                try
                                                {
                                                    if (DocUtils.CheckExits(pathServer + "/" + fileName))
                                                    {
                                                        DocUtils.DownloadFile(pathLocal, fileName, pathServer + "/" + fileName);
                                                    }
                                                }
                                                catch
                                                {                                   
                                                }
                                            }
                                        }
                        
                                    }
                                    catch
                                    {
                                    }
                                }
                                #endregion Mặt

                                #region TEM
                                if (fileModel.TEM)
                                {
                                    try
                                    {
                                        string[] listFileTEM = DocUtils.GetFilesList(pathServer);
                                        if (listFileTEM != null)
                                        {
                                            foreach (string fileName in listFileTEM)
                                            {
                                                try
                                                {
                                                    if (DocUtils.CheckExits(pathServer + "/" + fileName))
                                                    {
                                                        DocUtils.DownloadFile(pathLocal, fileName, pathServer + "/" + fileName);
                                                    }
                                                }
                                                catch
                                                {
                                                }
                                            }
                                        }

                                    }
                                    catch
                                    {
                                    }
                                }
                                #endregion TEM
                            }
                            else//1 file
                            {                                
                                if (strucModel.FolderType != 2)//nếu không phải là thư mục điện tử
                                {
                                    string filePathServer = fileModel.FolderContain.Replace("group", _productCode.Substring(0, 6)).Replace("code", _productCode);
                                    try
                                    {
                                        if (DocUtils.CheckExits(filePathServer))
                                        {
                                            DocUtils.DownloadFile(pathLocal, Path.GetFileName(filePathServer), filePathServer);
                                        }
                                    }
                                    catch
                                    {
                                    }
                                }
                                else
                                {
                                    DataRow[] listPCB = dtDMVT.Select("F4 like '%PCB%'");
                                    foreach (DataRow rowDT in listPCB)
                                    {
                                        string pcbCode = rowDT["F4"].ToString();
                                        string pathPCB = pathLocal + "/" + pcbCode;
                                        if (pathPCB == "") continue;
                                        Directory.CreateDirectory(pathPCB);

                                        string filePathServer = fileModel.FolderContain.Replace("tpat", "TPAT." + pcbCode.Substring(4, 7)).Replace("code", pcbCode);
                                        try
                                        {
                                            if (DocUtils.CheckExits(filePathServer))
                                            {
                                                DocUtils.DownloadFile(pathPCB, Path.GetFileName(filePathServer), filePathServer);
                                            }
                                        }
                                        catch
                                        {
                                        }
                                    }                                    
                                }
                            }
                        }
                    }
                }
               
            }

        }

        private void btnTestEncrypt_Click(object sender, EventArgs e)
        {
            frmEncrypt frm = new frmEncrypt();
            TextUtils.OpenForm(frm);
        }

        void SendMailGmail(string To, string Subject, string Body)
        {
            string smtpAddress = "smtp.gmail.com";
            int portNumber = 587;
            bool enableSSL = true;

            string emailFrom = "yourgmail";
            string password = "yourpass";
            string emailTo = To;
            string subject = Subject;
            string body = Body;

            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(emailFrom);
                mail.To.Add(emailTo);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true;
                // Can set to false, if you are sending pure text.

                //mail.Attachments.Add(new Attachment("C:\\SomeFile.txt"));
                //mail.Attachments.Add(new Attachment("C:\\SomeZip.zip"));

                SmtpClient smtp = new SmtpClient(smtpAddress, portNumber);
                smtp.Credentials = new NetworkCredential(emailFrom, password);
                smtp.EnableSsl = enableSSL;
                smtp.Send(mail);
            }
            catch
            {
                MessageBox.Show("Kết nối Internet của bạn đang có vấn đề!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Gửi google mail
        private void button11_Click(object sender, EventArgs e)
        {            
          
        }

        //download cad tổng,dmvt bản cứng
        private void button12_Click(object sender, EventArgs e)
        {
            DocUtils.InitFTPQLSX();
            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang download file..."))
            {
                ArrayList listModules = ModulesBO.Instance.FindByExpression(new Expression("Status", 2).And(new Expression("Code", "TPAD", "like")));
                string pathLocal = "D:/Ban Cad Tong, DMVT bản cứng";
                Directory.CreateDirectory(pathLocal);
                foreach (var item in listModules)
                {
                    ModulesModel module = (ModulesModel)item;
                    string code = module.Code;
                    string filePathCADCung = string.Format(@"Thietke.Ck\{0}\{1}.Ck\BCCk.{1}\BC-CAD.{1}\{1}.jpg", code.Substring(0, 6), code);
                    string filePathDMVTCung = string.Format(@"Thietke.Ck\{0}\{1}.Ck\BCCk.{1}\BC-VT.{1}\VT.{1}.pdf", code.Substring(0, 6), code);
                    if (DocUtils.CheckExits(filePathCADCung))
                    {
                        string modelPath = pathLocal + "/" + code;
                        Directory.CreateDirectory(modelPath);
                        try
                        {
                            DocUtils.DownloadFile(modelPath, code + ".jpg", filePathCADCung);
                        }
                        catch
                        {
                        }
                        try
                        {
                            DocUtils.DownloadFile(modelPath, "VT." + code + ".pdf", filePathDMVTCung);
                        }
                        catch
                        {
                        }
                    }
                    
                }
            }
        }
    }
}
