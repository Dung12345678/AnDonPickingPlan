using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using BMS.Model;
using BMS.Utils;
using System.IO;
using System.Collections;
using BMS.Business;
using System.Threading;
using System.Diagnostics;
using System.Data.SqlClient;
using KarlsTools;

namespace BMS
{
	public partial class frmLogin : Form
	{
		public bool loginSuccess = false;
		public string loginname = "";
		public string pass = "";
		string pImage = "";
		public frmLogin()
		{
			InitializeComponent();

			//CheckAutoUpdate();
		}
		private void frmLogin_Load(object sender, EventArgs e)
		{
			getLastlog();

			txtUserName.Focus();
			if (txtUserName.Text != "")
			{
				txtPassword.Select();
			}
			txtPassword.Text = "123456";
			//ProcessLogIn();
		}

		private void btnOK_Click(object sender, EventArgs e)
		{
			ProcessLogIn();
		}
		private void btnExit_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void ProcessLogIn()
		{

			//Global.DefaultFileName = "default.ini";

			#region Validate Login
			if (txtUserName.Text.Trim() == "" || txtPassword.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập đầy đủ tên đăng nhập và mật khẩu !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			#endregion

			#region Khai báo biến
			UsersModel userModel = null;
			string _user = txtUserName.Text;
			string _pass = txtPassword.Text;
			#endregion

			#region Nếu không phải thì kiểm tra
			Global.IsReal = !chkIsReal.Checked;
			IE.Global.IsReal = !chkIsReal.Checked;
			if (TextUtils.Log(_user, _pass, ref userModel) == true)
			{
				Global.UserID = userModel.ID;
				Global.LoginName = userModel.LoginName;
				Global.AppUserName = userModel.LoginName;
				Global.AppFullName = userModel.FullName;
				Global.AppPassword = _pass;
				//Global.MainViewID = userModel.MainViewID;
				Global.DepartmentID = userModel.DepartmentID;
				Global.UserGroupID = userModel.UserGroupID;
				Global.IsAdmin = userModel.IsAdmin;
			

				#region Create Session
				//Global.IsNotCreateSession = true;
				//Expression exp = new Expression("UserID", Global.UserID, "=");
				//exp = exp.And(new Expression("Status", 0, "="));
				//ArrayList arr = SessionBO.Instance.FindByExpression(exp);
				//if (arr.Count > 0)
				//{
				//    for (int i = 0; i < arr.Count; i++)
				//    {
				//        SessionModel _modelSession = (SessionModel)arr[i];
				//        _modelSession.Status = 2;
				//        _modelSession.EndTime = TextUtils.GetSystemDate();

				//        SessionBO.Instance.Update(_modelSession);
				//    }
				//}
				//SessionModel _sessionModel = new SessionModel();
				//_sessionModel.ComputerName = Global.ComputerName;
				//_sessionModel.UserID = Global.UserID;
				//_sessionModel.StartTime = TextUtils.GetSystemDate();
				//_sessionModel.EndTime = _sessionModel.StartTime;
				//_sessionModel.Status = 0;
				//Global.SessionID = Convert.ToInt32(SessionBO.Instance.Insert(_sessionModel));
				#endregion

				saveLaslog();

				loginSuccess = true;
			}
			else
			{
				MessageBox.Show(this, "Sai tên đăng nhập hoặc mật khẩu, vui lòng kiểm tra lại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
				txtPassword.SelectAll();
				loginSuccess = false;
				return;
			}
			this.Dispose();

			#endregion
		}

		private void frmLogin_KeyDown(object sender, KeyEventArgs e)
		{
			string key = "";
			key = key + e.KeyCode;
			if (key.Contains("ControlKeyN"))
			{
				key = "";
				//tsSave.Enabled = true;
				//clearForm();
			}

			if (key.Contains("F4") && key.Contains("Alt"))
			{
				Application.Exit();
			}
		}

		private void getLastlog()
		{
			string pLastLog = "";
			string pServerID = "";
			string pZone = "";
			//Read from file LastLog.ini
			string strPath = Application.StartupPath.ToString() + @"\LastLog.ini";
			FileStream file = new FileStream(strPath, FileMode.OpenOrCreate, FileAccess.Read);
			StreamReader sr = new StreamReader(file);
			int i = 0;
			while (i < 4)
			{
				if (i == 0)
					pLastLog = sr.ReadLine();
				if (i == 1)
					pServerID = sr.ReadLine();
				if (i == 2)
					pImage = sr.ReadLine();
				if (i == 3)
					pZone = sr.ReadLine();
				i = i + 1;
			}
			sr.Close();
			file.Close();
			txtUserName.Text = pLastLog;
		}
		private void saveLaslog()
		{
			bool isOK = true;
			string strPath;
			strPath = Application.StartupPath.ToString() + @"\LastLog.ini";
			File.Delete(strPath);
			FileStream file = new FileStream(strPath, FileMode.OpenOrCreate, FileAccess.Write);
			StreamWriter w = new StreamWriter(file);
			w.Write(txtUserName.Text);
			w.Write("\r\n");
			w.Write("default.ini");

			w.Write("\r\n");
			w.Write(pImage);
			w.Close();
			file.Close();
		}

		private void textBox2_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				ProcessLogIn();
			}
		}

		private void simpleButton3_Click(object sender, EventArgs e)
		{
			this.WindowState = FormWindowState.Minimized;
		}

		private void simpleButton2_Click(object sender, EventArgs e)
		{
			btnExit_Click(null, null);
		}

		private void groupControl1_Paint(object sender, PaintEventArgs e)
		{

		}

		private void panelControl1_Paint(object sender, PaintEventArgs e)
		{

		}

		private void panelControl1_Paint_1(object sender, PaintEventArgs e)
		{

		}

		private void panelControl1_Paint_2(object sender, PaintEventArgs e)
		{

		}
	}
}
