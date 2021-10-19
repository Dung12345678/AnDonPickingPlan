using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Business;
using BMS.Model;
using BMS.Utils;
using System.Collections;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace BMS
{
    public partial class frmShowExcel : _Forms
    {
        public frmShowExcel()
        {
            InitializeComponent();
        }

        void updateSungLuc()
        {
            List<string> lstError = new List<string>();
            for (int i = 0; i < grvData.RowCount; i++)
            {
                string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
                try
                {
                    //string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
                    string GunNumber = Lib.ToString(grvData.GetRowCellValue(i, "F23"));
                    string JobNumber = Lib.ToString(grvData.GetRowCellValue(i, "F24"));
                    string QtyOcBanGa = Lib.ToString(grvData.GetRowCellValue(i, "F27"));
                    string QtyOcBanThat = Lib.ToString(grvData.GetRowCellValue(i, "F28"));

                    string[] numberGuns = Regex.Split(GunNumber, @"[^0-9\.]+");
                    string[] numberJobNumbers = Regex.Split(JobNumber, @"[^0-9\.]+");

                    int gun = TextUtils.ToInt(numberGuns[1]);
                    int job = TextUtils.ToInt(numberJobNumbers[1]);

                    string sql = string.Format(@"UPDATE dbo.Product SET GunNumber = {0},JobNumber = {1}, QtyOcBanGa= {2}, QtyOcBanThat = {3} WHERE ProductCode ='{4}'"
                                            , gun
                                            , job
                                            , TextUtils.ToInt(QtyOcBanGa)
                                            , TextUtils.ToInt(QtyOcBanThat)
                                            , productCode);

                    TextUtils.ExcuteSQL(sql);
                }
                catch 
                {
                    //MessageBox.Show(productCode);
                }
            }
            MessageBox.Show("OK");
        }

        void CreateProduct()
        {
            DataTable dtGroup = Lib.Select("SELECT ID,ProductGroupCode FROM dbo.ProductGroup");
            List<BaseModel> lstProduct = ProductBO.Instance.GetAllListObject();
            //object[] lst = arrProduct.ToArray();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = grvData.RowCount - 1;
            List<string> lstError = new List<string>();
            char spt = '~';
            for (int i = 0; i < grvData.RowCount; i++)
            {
                progressBar1.Value = i;
                //Stopwatch stopwatch = Stopwatch.StartNew();
                if (i < 4) continue;

                string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
                string groupCode = Lib.ToString(grvData.GetRowCellValue(i, "F2"));
                //Kiểm tra nếu mã nhóm hoặc mã sản phẩm trống thì next
                if (string.IsNullOrEmpty(groupCode) || string.IsNullOrEmpty(productCode))
                {
                    continue;
                }
                try
                {
                    DataRow[] arrG = dtGroup.Select(string.Format("ProductGroupCode = '{0}'", groupCode));
                    int groupID = 0;
                    if (arrG.Length > 0)
                    {
                        groupID = Lib.ToInt(arrG[0]["ID"]);
                    }
                    else
                    {
                        continue;
                    }

                    ProductModel productModel = new ProductModel();
                    var lstObject = lstProduct.FirstOrDefault(o => ((ProductModel)o).ProductCode == productCode);
                    //ArrayList arrP = ProductBO.Instance.FindByAttribute("ProductCode", productCode);
                    if (lstObject != null)
                    {
                        productModel = (ProductModel)lstObject;
                    }

                    productModel.ProductGroupID = groupID;
                    productModel.ProductGroupCode = groupCode;
                    productModel.ProductCode = productCode;
                    productModel.ProductName = Lib.ToString(grvData.GetRowCellValue(i, "F3"));//Mô tả-describe
                    productModel.ProductTypeCode = Lib.ToString(grvData.GetRowCellValue(i, "F4"));//Mô tả-describe

                    productModel.RatioCode = Lib.ToString(grvData.GetRowCellValue(i, "F5"));//Ratio
                    productModel.PinNumber = Lib.ToString(grvData.GetRowCellValue(i, "F6"));//SỐ PIN NGOÀI
                    productModel.PinNumberIN = Lib.ToString(grvData.GetRowCellValue(i, "F7"));//SỐ PIN TRONG

                    string valueString = Lib.ToString(grvData.GetRowCellValue(i, "F8"));

                    productModel.SizeSSSMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//KÍCH THƯỚC CỤM TRỤC SSS
                    productModel.SizeSSSMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//KÍCH THƯỚC CỤM TRỤC SSS
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F9"));
                    productModel.DoDao1Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 1
                    productModel.DoDao1Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 1
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F10"));
                    productModel.DoDao2Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 2
                    productModel.DoDao2Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 2
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F11"));
                    productModel.DoDao3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 3
                    productModel.DoDao3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 3
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F12"));
                    productModel.DoDao4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ ĐẢO 4
                    productModel.DoDao4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ ĐẢO 4

                    productModel.MotorCode = Lib.ToString(grvData.GetRowCellValue(i, "F13"));//MÃ MOTOR
                    productModel.LoaiMo = Lib.ToString(grvData.GetRowCellValue(i, "F14"));//LOẠI MỠ
                    productModel.LuongMo = Lib.ToDecimal(grvData.GetRowCellValue(i, "F15"));//LƯỢNG MỠ TỔNG

                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F16"));
                    productModel.LuongMo1Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho trục tốc độ thấp
                    productModel.LuongMo1Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho trục tốc độ thấp
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F17"));
                    productModel.LuongMo2Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lượng mỡ cho hộp
                    productModel.LuongMo2Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lượng mỡ cho hộp
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F18"));
                    productModel.LuongMo3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho vòng bi chịu tải & Carrier SSS
                    productModel.LuongMo3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho vòng bi chịu tải & Carrier SSS
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F19"));
                    productModel.LuongMo4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//LƯƠNG MỠ CHO VÒNG BI KHÔNG TẢI
                    productModel.LuongMo4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//LƯƠNG MỠ CHO VÒNG BI KHÔNG TẢI
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F20"));
                    productModel.LuongMo5Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho Carrier
                    productModel.LuongMo5Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho Carrier
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F21"));
                    productModel.LuongMo6Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512
                    productModel.LuongMo6Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Mỡ cho hộp (Gram) ở công đoạn 3, số thứ tự 50 của 511,512

                    productModel.MEDEDAY = Lib.ToDate2(grvData.GetRowCellValue(i, "F22"));//DATE
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F23"));
                    productModel.LucCheckGearMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check gear Unit
                    productModel.LucCheckGearMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check gear Unit

                    productModel.GunNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F24"));//Súng
                    productModel.JobNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F25"));//JOB
                    productModel.BuLongNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F26"));//SỐ LUỢNG BULONG
                    productModel.ShootNumber = Lib.ToInt(grvData.GetRowCellValue(i, "F27"));//SỐ LẦN BẮN
                    productModel.QtyOcBanGa = Lib.ToInt(grvData.GetRowCellValue(i, "F28"));//BẮN GÁ
                    productModel.SetForce = Lib.ToInt(grvData.GetRowCellValue(i, "F29"));//ĐẶT LỰC

                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F30"));
                    productModel.LucCheckGearMotorMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check Gear Motor
                    productModel.LucCheckGearMotorMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check Gear Motor

                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F31"));
                    productModel.LucCheck3Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check cụm trục tốc độ thấp 511-512
                    productModel.LucCheck3Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check cụm trục tốc độ thấp 511-512
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F32"));
                    productModel.LucCheck4Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check phần xuất lực Naknishi
                    productModel.LucCheck4Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check phần xuất lực Naknishi
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F33"));
                    productModel.LucCheck5Min = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Lực check Ốc chặn mỡ Nakanishi, 6000
                    productModel.LucCheck5Max = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Lực check Ốc chặn mỡ Nakanishi, 6000

                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F34"));
                    productModel.SpaceCenterMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//KHOẢNG CÁCH TRỤC TRUNG TÂM.
                    productModel.SpaceCenterMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//KHOẢNG CÁCH TRỤC TRUNG TÂM.

                    productModel.MURATA = Lib.ToString(grvData.GetRowCellValue(i, "F35"));//MURATA
                    productModel.CoilCode = Lib.ToString(grvData.GetRowCellValue(i, "F36"));//MÃ COIL
                    productModel.DienAp = Lib.ToDecimal(grvData.GetRowCellValue(i, "F37"));//Điện áp (V)
                    productModel.TanSo = Lib.ToDecimal(grvData.GetRowCellValue(i, "F38"));//Tần số

                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F39"));
                    productModel.DongDienMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Dòng điện Io
                    productModel.DongDienMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Dòng điện Io
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F40"));
                    productModel.NhapLucMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Nhập lực Wo
                    productModel.NhapLucMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Nhập lực Wo
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F41"));
                    productModel.VongQuayMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Vòng quay
                    productModel.VongQuayMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Vòng quay
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F42"));
                    productModel.TiengOnMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//TIẾNG ỒN
                    productModel.TiengOnMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//TIẾNG ỒN
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F43"));
                    productModel.DoRungFFTMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//ĐỘ RUNG FFT
                    productModel.DoRungFFTMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//ĐỘ RUNG FFT
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F44"));
                    productModel.SpaceCenterMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//khoảng cách A giữa hai mặt vòng bi Cycloid disk
                    productModel.SpaceCenterMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//khoảng cách A giữa hai mặt vòng bi Cycloid disk
                    valueString = Lib.ToString(grvData.GetRowCellValue(i, "F45"));
                    productModel.DifferenceMin = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[0] : "0");//Độ chênh lệch lớn nhất tại 3 vị trí đo 0.03mm
                    productModel.DifferenceMax = Lib.ToDecimal(valueString.Contains(spt) ? valueString.Split(spt)[1] : "0");//Độ chênh lệch lớn nhất tại 3 vị trí đo 0.03mm

                    List<string> lstCommand = new List<string>();

                    if (productModel.ID == 0)
                    {
                        productModel.ID = (int)ProductBO.Instance.Insert(productModel);
                        //Thêm các hạng mục check theo nhóm sản phẩm
                        string sqlInitWorking = string.Format("EXEC dbo.CreateProductWorking_ByGroupID {0},{1}", groupID, productModel.ID);
                        lstCommand.Add(sqlInitWorking);
                        // TextUtils.ExcuteSQL(sqlInitWorking);
                        lstProduct.Add(productModel);
                    }
                    else
                    {
                        ProductBO.Instance.Update(productModel);
                    }
                    
                    //Sinh các câu lệnh update động các mục kiểm tra
                    for (int j = 0; j < grvData.Columns.Count; j++)
                    {
                        string rule = Lib.ToString(grvData.GetRowCellValue(3, grvData.Columns[j]));
                        if (string.IsNullOrEmpty(rule)) continue;
                        string value = Lib.ToString(grvData.GetRowCellValue(i, grvData.Columns[j]));
                        if (string.IsNullOrEmpty(value)) continue;

                        string min = "";
                        string max = "";
                        if (value.Contains(spt))
                        {
                            min = value.Split(spt)[0].Trim();
                            max = value.Split(spt)[1].Trim();
                        }
                        else
                        {
                            min = max = Lib.ToDecimal(value).ToString();
                        }
                        
                        string[] arrCD = rule.Split(';');
                        for (int h = 0; h < arrCD.Length; h++)
                        {
                            string[] arrW = arrCD[h].Trim().Split('_');
                            if (arrW.Length <= 1) continue;
                            string stepCode = arrW[0].Trim();
                            
                            string sortOrder = arrW[1].Trim();
                            string sql = getSql(value, min, max, sortOrder, stepCode, productModel.ID);
                            lstCommand.Add(sql);
                        }
                    }

                    string sqlUpdateWorking = string.Format("EXEC spUpdateProductWorking_ByProductID_New {0},{1}", productModel.ID, productModel.MURATA == "3" ? 0 : 1);
                    lstCommand.Add(sqlUpdateWorking);
                    //TextUtils.ExcuteSQL(sqlUpdateWorking);

                    if (lstCommand.Count > 0)
                    {
                        Lib.ExcuteSQL(string.Join(";\n", lstCommand));
                    }
                }
                catch(Exception ex)
                {
                    lstError.Add(productCode);
                }
                //stopwatch.Stop();
                //MessageBox.Show(stopwatch.ElapsedMilliseconds.ToString());
            }
            //MessageBox.Show(DateTime.Now.ToString());
        }

        string getSql(string periodValue, string min, string max, string sortOrder, string stepCode, int productID)
        {
            string sql = "";
            sql = string.Format(@"UPDATE w 
	                            SET  w.PeriodValue = '{0}', w.MinValue = {1}, w.MaxValue = {2}
	                            FROM dbo.ProductWorking w
		                            JOIN dbo.ProductStep s ON s.ID = w.ProductStepID AND s.ProductStepCode = '{3}'
	                            WHERE w.SortOrder IN ({4})  AND w.ProductID = {5}", periodValue, min, max, stepCode, sortOrder, productID);
            return sql;
        }
        
        void creatWorkingStep()
        {
            for (int i = 1; i < 6; i++)
            {
                for (int j = 1; j < grvData.RowCount; j++)
                {
                    WorkingStepModel gStep = new WorkingStepModel();
                    gStep.ProductGroupID = i;
                    gStep.Description = Lib.ToString(grvData.GetRowCellValue(j, "F3")).Trim();
                    gStep.SortOrder = Lib.ToInt(grvData.GetRowCellValue(j, "F1"));
                    gStep.WorkingStepCode = "CD"+ Lib.ToString(grvData.GetRowCellValue(j, "F2")).Trim();

                    gStep.ID = (int)WorkingStepBO.Instance.Insert(gStep);
                }

                Lib.ExcuteSQL(string.Format("EXEC CreateProductStep_ByGroupID {0}", i));
            }
        }

        void createStep1()
        {
            for (int i = 0; i < grvData.Columns.Count; i++)
            {
                if (i < 5)
                {
                    continue;
                }
                string workingID = Lib.ToString(grvData.GetRowCellValue(1, grvData.Columns[i]));

                for (int j = 0; j < grvData.RowCount; j++)
                {
                    string groupCode = Lib.ToString(grvData.GetRowCellValue(j, "F2"));
                    if (string.IsNullOrWhiteSpace(groupCode)) continue;

                    string productCode = Lib.ToString(grvData.GetRowCellValue(j, "F1"));
                    string value = Lib.ToString(grvData.GetRowCellValue(j, grvData.Columns[i]));
                    
                    string sqlUpdate = @"UPDATE PW 
                                SET PW.PeriodValue = '{0}'
                                FROM dbo.ProductWorking PW JOIN dbo.Product P ON P.ID = PW.ProductID
                                WHERE PW.WorkingID = {1} AND P.ProductCode = N'{2}'";
                    Lib.ExcuteSQL(string.Format(sqlUpdate, value, workingID, productCode));
                }
            }
        }

        void UpdateDoDao()
        {
            for (int i = 0; i < grvData.RowCount; i++)
            {
                string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
                string dodao = Lib.ToString(grvData.GetRowCellValue(i, "F20"));
                string sql = string.Format(@"UPDATE dbo.Product SET Loai2TrucCoChan = '{0}' WHERE ProductCode = '{1}'", dodao, productCode);
                Lib.ExcuteSQL(sql);
            }
        }

        void updateValueCD1()
        {
            for (int i = 0; i < grvData.RowCount; i++)
            {
                string productCode = Lib.ToString(grvData.GetRowCellValue(i, "F1"));
                if (string.IsNullOrEmpty(productCode))
                {
                    continue;
                }
                string v10 = Lib.ToString(grvData.GetRowCellValue(i, "F6"));
                string v20 = Lib.ToString(grvData.GetRowCellValue(i, "F7"));
                string v30 = Lib.ToString(grvData.GetRowCellValue(i, "F8"));
                string v40 = Lib.ToString(grvData.GetRowCellValue(i, "F9"));
                string v50 = Lib.ToString(grvData.GetRowCellValue(i, "F10"));
                string v60 = Lib.ToString(grvData.GetRowCellValue(i, "F11"));
                string v70 = Lib.ToString(grvData.GetRowCellValue(i, "F12"));
                string v80 = Lib.ToString(grvData.GetRowCellValue(i, "F13"));
                string v90 = Lib.ToString(grvData.GetRowCellValue(i, "F14"));
                string v100 = Lib.ToString(grvData.GetRowCellValue(i, "F15"));
                string v110 = Lib.ToString(grvData.GetRowCellValue(i, "F16"));
                string v120 = Lib.ToString(grvData.GetRowCellValue(i, "F17"));

                TextUtils.ExcuteProcedure("spImportCD1", new string[] {
                    "@ProductCode","@Value10" ,"@Value20" ,"@Value30" ,"@Value40" ,"@Value50" ,"@Value60" ,
	                "@Value70" ,"@Value80" ,"@Value90" ,"@Value100","@Value110","@Value120"
                }, new object[] { productCode,v10, v20, v30, v40, v50, v60, v70, v80, v90, v100, v110, v120 });
            }
            MessageBox.Show("Nhập khẩu thành công dữ liệu công đoạn 1");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //updateValueCD1();
            CreateProduct();
            //creatWorkingStep();
            //createStep1();
            //UpdateDoDao();

            //updateSungLuc();
            //splitGearCheck();
        }

        void splitGearCheck()
        {   
            ArrayList lst = ProductBO.Instance.FindAll();
            for (int i = 0; i < lst.Count; i++)
            {
                ProductModel model = (ProductModel)lst[i];
                string b = model.LucCheckGear;
                if (!string.IsNullOrEmpty(b))
                {
                    string[] numbers = Regex.Split(b, @"[^0-9\.]+");
                    if (numbers.Length > 1)
                    {
                        model.LucCheckGearMin = TextUtils.ToDecimal(numbers[0]);
                        model.LucCheckGearMax = TextUtils.ToDecimal(numbers[1]);
                    }
                }

                b = model.LucCheckGearMotor;
                if (!string.IsNullOrEmpty(b))
                {
                    string[] numbers = Regex.Split(b, @"[^0-9\.]+");
                    if (numbers.Length > 1)
                    {
                        model.LucCheckGearMotorMin = TextUtils.ToDecimal(numbers[0]);
                        model.LucCheckGearMotorMax = TextUtils.ToDecimal(numbers[1]);
                    }
                }
                
                ProductBO.Instance.Update(model);
            }
        }

        private void btnBrowse_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                btnBrowse.Text = ofd.FileName;
                cboSheet.DataSource = null;
                cboSheet.DataSource = TextUtils.ListSheetInExcel(ofd.FileName);
            }
        }

        private void cboSheet_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = TextUtils.ExcelToDatatableNoHeader(btnBrowse.Text, cboSheet.SelectedValue.ToString());

                grdData.DataSource = dt;
                grvData.PopulateColumns();
                grvData.BestFitColumns();
            }
            catch (Exception ex)
            {
                TextUtils.ShowError(ex);
                grdData.DataSource = null;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            TextUtils.ExportExcel(grvData);
        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
