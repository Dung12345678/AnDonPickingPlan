using BMS.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public static class Global
	{
		//static string _ConnectionString = @"Server=LAPTOP-G3GTTN00\SQLEXPRESS;Database=SumitomoTest;uid=sa;pwd=123@123a;";
			static string _ConnectionString = "";     
		public static string ConnectionString
		{
			get
			{
				if (string.IsNullOrEmpty(_ConnectionString))
				{
					string strPath = Application.StartupPath.ToString() + (IsReal ? @"\default.ini" : @"\defaultBackup.ini");
					_ConnectionString = MD5.DecodeChecksum(File.ReadAllText(strPath));
				}
				return _ConnectionString;
			}
			set { _ConnectionString = value; }
		}

		public static bool IsReal { get; set; } = true;

		static int _AppUserID;
		static int _AppGroupID;
		static string _AppUserName = "";
		static string _AppFullName = "";
		static string _AppPassword;
		static int _AppRegionID;
		static bool _IsRoot = false;
		static string _AppUserCode;
		public static int DepartmentID = 1;
		static int _ShiftID = 0;
		static string _ServerName = "";
		static string _DatabaseName = "";
		static int _OutletID;
		static int _ZoneID;
		static int _ComputerID;

		public static string ComputerName { get; set; }
		public static string ComputerMACAddresses { get; set; }
		public static bool IsTest { get; set; }

		/// <summary>
		/// Access routine for global variable.
		/// </summary>
		public static int ShiftID
		{
			get { return _ShiftID; }
			set { _ShiftID = value; }
		}


		/// <summary>
		/// Access routine for global variable.
		/// </summary>
		public static string ServerName
		{
			get { return _ServerName; }
			set { _ServerName = value; }
		}
		/// <summary>
		/// Access routine for global variable.
		/// </summary>
		public static string DatabaseName
		{
			get { return _DatabaseName; }
			set { _DatabaseName = value; }
		}

		public static string LoginName
		{
			get { return Global._AppUserCode; }
			set { Global._AppUserCode = value; }
		}

		public static int UserID
		{
			get { return _AppUserID; }
			set { _AppUserID = value; }
		}



		public static string AppUserName
		{
			get { return _AppUserName; }
			set { _AppUserName = value; }
		}

		public static string AppFullName
		{
			get { return _AppFullName; }
			set { _AppFullName = value; }
		}

		public static string AppPassword
		{
			get { return _AppPassword; }
			set { _AppPassword = value; }
		}
		public static string DefaultFileName { get; set; }

		public static int UserGroupID { get; set; }
		public static bool IsAdmin { get; set; }
	}

	public enum EnumConnectionStringIndex : int
	{
		Altax = 0,
		Hyp = 1,
	}
}
