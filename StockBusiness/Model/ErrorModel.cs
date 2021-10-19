
using System;
namespace ST.Model
{
	public class ErrorModel : BaseModel
	{
		private int iD;
		private string name;
		private string code;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string Name
		{
			get { return name; }
			set { name = value; }
		}
	
		public string Code
		{
			get { return code; }
			set { code = value; }
		}
	
	}
}
	