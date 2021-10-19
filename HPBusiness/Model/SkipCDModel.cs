
using System;
namespace HP.Model
{
	public class SkipCDModel : BaseModel
	{
		private int iD;
		private string product;
		private string cDSkip;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string Product
		{
			get { return product; }
			set { product = value; }
		}
	
		public string CDSkip
		{
			get { return cDSkip; }
			set { cDSkip = value; }
		}
	
	}
}
	