
using System;
namespace HP.Model
{
	public class CheckHistoryDetailModel : BaseModel
	{
		private int iD;
		private int checkHistoryID;
		private string workingName;
		private int checkValueType;
		private string unit;
		private string standardValue;
		private string realValue;
		private int sortOrder;
		private string result;
		private DateTime? createDate;
		private int isDisplay;
		private int x;
		private int y;
		private decimal min;
		private decimal max;
		private string realValue1;
		private string realValue2;
		private string realValue3;
		private string realValue4;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int CheckHistoryID
		{
			get { return checkHistoryID; }
			set { checkHistoryID = value; }
		}
	
		public string WorkingName
		{
			get { return workingName; }
			set { workingName = value; }
		}
	
		public int CheckValueType
		{
			get { return checkValueType; }
			set { checkValueType = value; }
		}
	
		public string Unit
		{
			get { return unit; }
			set { unit = value; }
		}
	
		public string StandardValue
		{
			get { return standardValue; }
			set { standardValue = value; }
		}
	
		public string RealValue
		{
			get { return realValue; }
			set { realValue = value; }
		}
	
		public int SortOrder
		{
			get { return sortOrder; }
			set { sortOrder = value; }
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
	
		public int IsDisplay
		{
			get { return isDisplay; }
			set { isDisplay = value; }
		}
	
		public int X
		{
			get { return x; }
			set { x = value; }
		}
	
		public int Y
		{
			get { return y; }
			set { y = value; }
		}
	
		public decimal Min
		{
			get { return min; }
			set { min = value; }
		}
	
		public decimal Max
		{
			get { return max; }
			set { max = value; }
		}
	
		public string RealValue1
		{
			get { return realValue1; }
			set { realValue1 = value; }
		}
	
		public string RealValue2
		{
			get { return realValue2; }
			set { realValue2 = value; }
		}
	
		public string RealValue3
		{
			get { return realValue3; }
			set { realValue3 = value; }
		}
	
		public string RealValue4
		{
			get { return realValue4; }
			set { realValue4 = value; }
		}
	
	}
}
	