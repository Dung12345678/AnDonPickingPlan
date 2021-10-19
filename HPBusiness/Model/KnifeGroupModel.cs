
using System;
namespace HP.Model
{
	public class KnifeGroupModel : BaseModel
	{
		private int iD;
		private string knifeGroupCode;
		private string knifeGroupName;
		private DateTime? createdDate;
		private int type;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string KnifeGroupCode
		{
			get { return knifeGroupCode; }
			set { knifeGroupCode = value; }
		}
	
		public string KnifeGroupName
		{
			get { return knifeGroupName; }
			set { knifeGroupName = value; }
		}
	
		public DateTime? CreatedDate
		{
			get { return createdDate; }
			set { createdDate = value; }
		}
	
		public int Type
		{
			get { return type; }
			set { type = value; }
		}
	
	}
}
	