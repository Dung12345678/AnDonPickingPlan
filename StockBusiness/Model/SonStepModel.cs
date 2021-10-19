
using System;
namespace ST.Model
{
	public class SonStepModel : BaseModel
	{
		private int iD;
		private string sonStepCode;
		private string sonStepName;
		private DateTime? createDate;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string SonStepCode
		{
			get { return sonStepCode; }
			set { sonStepCode = value; }
		}
	
		public string SonStepName
		{
			get { return sonStepName; }
			set { sonStepName = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
	}
}
	