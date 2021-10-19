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
using HP.Business;
using HP.Model;
using HP.Utils;
using DevExpress.Data.Filtering.Helpers;
using ExcelDataReader;
using System.Diagnostics;
using System.IO;
using System.Data.OleDb;
using System.Globalization;
using System.Text.RegularExpressions;

namespace BMS
{

	public partial class frmImportKnifeCount : _Forms
	{
		string path = "";
		DataTable dtt = new DataTable();
		public frmImportKnifeCount()
		{
			InitializeComponent();
		}
		private void frmUpdateDate_Load(object sender, EventArgs e)
		{

		}
		void UpdateDate()
		{
			List<string> lst = new List<string>();
			List<string> lstInt = new List<string>();
			string Order1 = "";
			string Order2 = "";
			if (Path.GetExtension(btnBrowse.Text).ToUpper() == ".TXT")
			{
				string filename = btnBrowse.Text;
				//Tạo bảng
				dtt = new DataTable();
				//THêm cột vào bảng
				dtt.Columns.Add("F1");
				dtt.Columns.Add("F2");
				dtt.Columns.Add("F3");
				dtt.Columns.Add("F4");
				dtt.Columns.Add("F5");
				dtt.Columns.Add("F6");
				dtt.Columns.Add("F7");
				dtt.Columns.Add("F8");
				dtt.Columns.Add("F9");
				dtt.Columns.Add("F10");
				dtt.Columns.Add("F11");
				dtt.Columns.Add("F12");
				dtt.Columns.Add("F13");
				dtt.Columns.Add("F14");
				dtt.Columns.Add("F15");
				dtt.Columns.Add("F16");
				//gọi hàm đọc file txt
				string noidung = Lib.GetFileContentTXT(filename);
				//Cắt xuống dòng -"\n"
				string[] strContent = noidung.Split('\n');
				foreach (string dong in strContent)
				{

					if (String.IsNullOrEmpty(dong))
						break;

					//Cắt dấu "|"
					string[] _dong = dong.Split('\t');

					//add các dồng đã cắt dấu "|" vào các dòng của bảng
					if (TextUtils.ToInt(_dong[6]) == 1 || TextUtils.ToInt(_dong[6]) == 2)
					{
						string[] AddDong = _dong[15].Split('|');

						if (TextUtils.ToInt(_dong[6]) == 1)
						{
							if (lst.Count > 0)
							{
								lst.Clear();
							}
							Order1 = _dong[1];
							for (int i = 0; i < AddDong.Count(); i++)
							{
								lst.Add(AddDong[i]);
							}
						}
						else if (TextUtils.ToInt(_dong[6]) == 2)
						{
							if (lstInt.Count > 0)
							{
								lstInt.Clear();
							}
							Order2 = _dong[1];
							for (int i = 0; i < AddDong.Count(); i++)
							{
								lstInt.Add(AddDong[i]);
							}
						}

					}
					else
					{
						continue;
					}
					if (lst.Count > 0 && lstInt.Count > 0 && (Order1.Trim() == Order2.Trim()))
					{
						for (int i = 0; i < lst.Count; i++)
						{
							DataRow dr1 = dtt.NewRow();
							dr1["F1"] = _dong[0];//Br
							dr1["F2"] = _dong[1];//OrderMachining
							dr1["F3"] = _dong[2];
							dr1["F4"] = _dong[3];//ProductCode
							dr1["F5"] = _dong[4];
							dr1["F6"] = _dong[5];
							dr1["F7"] = _dong[6];
							dr1["F8"] = _dong[7];//CreateDate
							dr1["F9"] = _dong[8];
							dr1["F10"] = _dong[9];
							dr1["F11"] = _dong[10];
							dr1["F12"] = _dong[11];
							dr1["F13"] = _dong[12];
							dr1["F14"] = _dong[13];
							dr1["F15"] = _dong[14];
							if (lstInt.Count < i + 1)
							{
								break;
							}
							string so = Regex.Replace(lstInt[i], ",", string.Empty).Trim();
							dr1["F16"] = lst[i].Trim() + so;//KnifeCode
							if (dr1["F16"].ToString().Trim() == "") continue;
							dtt.Rows.Add(dr1);
						}
						lst.Clear();
						lstInt.Clear();
						Order1 = "";
						Order2 = "";
					}
				}
				grdData.DataSource = null;
				grdData.DataSource = dtt;
			}
		}
		void Save()
		{
			int rowCount = dtt.Rows.Count;
			for (int i = 0; i < rowCount; i++)
			{
				try
				{
					string _KnifeCode = Lib.ToString(dtt.Rows[i]["F16"]);
					if (string.IsNullOrEmpty(_KnifeCode))
					{
						continue;
					}


					ProductKnifeModel productKnifeModel = new ProductKnifeModel();

					#region SetValue
					//string a = "";
					productKnifeModel.Code = _KnifeCode;
					productKnifeModel.BR = Lib.ToString(dtt.Rows[i]["F1"]);
					productKnifeModel.OrderMachining = Lib.ToString(dtt.Rows[i]["F2"]);
					productKnifeModel.GoodsCode = Lib.ToString(dtt.Rows[i]["F4"]);
					productKnifeModel.CreateDate = Lib.ToDate3(dtt.Rows[i]["F8"]);//date
					#endregion
					Expression exp1 = new Expression("KnifeCode", _KnifeCode);
					ArrayList arr = ProductKnifeBO.Instance.FindByExpression(exp1);
					if (arr.Count == 0)
					{
						ProductKnifeBO.Instance.Insert(productKnifeModel);
					}

				}
				catch (Exception ex)
				{
					//MessageBox.Show("Lỗi lưu dữ liệu tại dòng " + i + Environment.NewLine + ex.ToString());
				}
			}
		}
		private void btnBrowse_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
		{
			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			var result = openFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				btnBrowse.Text = openFileDialog1.FileName;
				path = btnBrowse.Text.Trim();
				UpdateDate();
				//File.WriteAllText(Application.StartupPath + "/UpdateDate.txt", path);
			}

		}

		public static string GetFileContent(string filepath)
		{
			string content = "";

			using (FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read))
			{
				using (StreamReader reader = new StreamReader(stream))
				{
					string line;
					while (reader.Peek() >= 0)
					{
						line = reader.ReadLine() + "\r\n";

						content += line;
					}
				}
			}
			return content;
		}

		//private void chkRun_CheckedChanged(object sender, EventArgs e)
		//{
		//	string path_my_app = Application.StartupPath + "\\RTC.exe";
		//	if (chkRun.Checked)
		//	{
		//		if (File.Exists(path_my_app))
		//		{
		//			SetStartup(path_my_app, true);
		//		}
		//	}
		//	else
		//	{
		//		if (File.Exists(path_my_app))
		//		{
		//			SetStartup(path_my_app, false);
		//		}
		//	}

		//}
		////auto chạy khi khởi động windowns
		//private void SetStartup(string AppName, bool enable)
		//{
		//	string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
		//	Microsoft.Win32.RegistryKey startupKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(runKey);
		//	if (enable)
		//	{
		//		if (startupKey.GetValue(AppName) == null)
		//		{
		//			startupKey.Close();
		//			startupKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(runKey, true);
		//			//startupKey.SetValue(AppName, Assembly.GetExecutingAssembly().Location + " /StartMinimized");
		//			startupKey.SetValue(AppName, Application.StartupPath + "\\RTC.exe");
		//			startupKey.Close();
		//		}
		//	}
		//	else
		//	{
		//		startupKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(runKey, true);
		//		startupKey.DeleteValue(AppName, false);
		//		startupKey.Close();
		//	}
		//}

		private void btnSave_Click(object sender, EventArgs e)
		{
			Save();
		}
	}
}
