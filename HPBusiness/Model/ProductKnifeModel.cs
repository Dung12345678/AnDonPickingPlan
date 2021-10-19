
using System;
namespace HP.Model
{
	public class ProductKnifeModel : BaseModel
	{
		private int iD;
		private string orderMachining;
		private string code;
		private string goodsCode;
		private DateTime? createDate;
		private string bR;
		private int quantity;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string OrderMachining
		{
			get { return orderMachining; }
			set { orderMachining = value; }
		}
	
		public string Code
		{
			get { return code; }
			set { code = value; }
		}
	
		public string GoodsCode
		{
			get { return goodsCode; }
			set { goodsCode = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public string BR
		{
			get { return bR; }
			set { bR = value; }
		}
	
		public int Quantity
		{
			get { return quantity; }
			set { quantity = value; }
		}
	
	}
}
	