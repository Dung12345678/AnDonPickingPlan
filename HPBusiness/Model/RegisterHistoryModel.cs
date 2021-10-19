
using System;
namespace HP.Model
{
	public class RegisterHistoryModel : BaseModel
	{
		private int iD;
		private int goodsID;
		private int toolsID;
		private DateTime? createDate;
		private bool isCheck;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int GoodsID
		{
			get { return goodsID; }
			set { goodsID = value; }
		}
	
		public int ToolsID
		{
			get { return toolsID; }
			set { toolsID = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public bool IsCheck
		{
			get { return isCheck; }
			set { isCheck = value; }
		}
	
	}
}
	