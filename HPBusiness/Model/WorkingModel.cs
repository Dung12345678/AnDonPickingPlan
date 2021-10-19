
using System;
namespace HP.Model
{
	public partial class WorkingModel : BaseModel
	{
		private int iD;
		private int workingStepID;
		private string workingName;
		private int valueType;
		private string valueTypeName;
		private string periodValue;
		private double minValue;
		private double maxValue;
		private int sortOrder;
		private DateTime? createdDate;
		private int checkValueType;
		private string unit;
		private int x;
		private int y;
		private int isDisplay;
		private string condition;
		private string frequency;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int WorkingStepID
		{
			get { return workingStepID; }
			set { workingStepID = value; }
		}
	
		public string WorkingName
		{
			get { return workingName; }
			set { workingName = value; }
		}
	
		public int ValueType
		{
			get { return valueType; }
			set { valueType = value; }
		}
	
		public string ValueTypeName
		{
			get { return valueTypeName; }
			set { valueTypeName = value; }
		}
	
		public string PeriodValue
		{
			get { return periodValue; }
			set { periodValue = value; }
		}
	
		public double MinValue
		{
			get { return minValue; }
			set { minValue = value; }
		}
	
		public double MaxValue
		{
			get { return maxValue; }
			set { maxValue = value; }
		}
	
		public int SortOrder
		{
			get { return sortOrder; }
			set { sortOrder = value; }
		}
	
		public DateTime? CreatedDate
		{
			get { return createdDate; }
			set { createdDate = value; }
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
	
		public int IsDisplay
		{
			get { return isDisplay; }
			set { isDisplay = value; }
		}
	
		public string Condition
		{
			get { return condition; }
			set { condition = value; }
		}
	
		public string Frequency
		{
			get { return frequency; }
			set { frequency = value; }
		}
	
	}
}
	