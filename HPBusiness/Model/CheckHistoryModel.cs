
using System;
namespace HP.Model
{
	public class CheckHistoryModel : BaseModel
	{
		private int iD;
		private int knifeDetailListID;
		private string orderCode;
		private DateTime? createDate;
		private int type;
		private string knifeDetailCode;
		private string machine;
		private string stepCode;
		private string workerCode;
		private string goodsCode;
		private int partID;
		private string knifeDetailCodeReal;
		private string partCode;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int KnifeDetailListID
		{
			get { return knifeDetailListID; }
			set { knifeDetailListID = value; }
		}
	
		public string OrderCode
		{
			get { return orderCode; }
			set { orderCode = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public int Type
		{
			get { return type; }
			set { type = value; }
		}
	
		public string KnifeDetailCode
		{
			get { return knifeDetailCode; }
			set { knifeDetailCode = value; }
		}
	
		public string Machine
		{
			get { return machine; }
			set { machine = value; }
		}
	
		public string StepCode
		{
			get { return stepCode; }
			set { stepCode = value; }
		}
	
		public string WorkerCode
		{
			get { return workerCode; }
			set { workerCode = value; }
		}
	
		public string GoodsCode
		{
			get { return goodsCode; }
			set { goodsCode = value; }
		}
	
		public int PartID
		{
			get { return partID; }
			set { partID = value; }
		}
	
		public string KnifeDetailCodeReal
		{
			get { return knifeDetailCodeReal; }
			set { knifeDetailCodeReal = value; }
		}
	
		public string PartCode
		{
			get { return partCode; }
			set { partCode = value; }
		}
	
	}
}
	