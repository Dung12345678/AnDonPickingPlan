
using System;
namespace HP.Model
{
	public class PartModel : BaseModel
	{
		private int iD;
		private string partCode;
		private string partName;
		private DateTime? createDate;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string PartCode
		{
			get { return partCode; }
			set { partCode = value; }
		}
	
		public string PartName
		{
			get { return partName; }
			set { partName = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
	}
}
	