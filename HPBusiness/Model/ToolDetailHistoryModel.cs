
using System;
namespace HP.Model
{
	public class ToolDetailHistoryModel : BaseModel
	{
		private int iD;
		private string checker;
		private int toolID;
		private double value;
		private string result;
		private DateTime? createDate;
		private double min;
		private double max;
		private double std;
		private string code;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string Checker
		{
			get { return checker; }
			set { checker = value; }
		}
	
		public int ToolID
		{
			get { return toolID; }
			set { toolID = value; }
		}
	
		public double Value
		{
			get { return value; }
			set { value = value; }
		}
	
		public string Result
		{
			get { return result; }
			set { result = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public double Min
		{
			get { return min; }
			set { min = value; }
		}
	
		public double Max
		{
			get { return max; }
			set { max = value; }
		}
	
		public double Std
		{
			get { return std; }
			set { std = value; }
		}
	
		public string Code
		{
			get { return code; }
			set { code = value; }
		}
	
	}
}
	