
using System;
namespace HP.Model
{
	public class KnifeJigWorkingModel : BaseModel
	{
		private int iD;
		private int knifeDetailListID;
		private int knifeStepID;
		private int workingID;
		private string workingName;
		private int valueType;
		private string valueTypeName;
		private string periodValue;
		private decimal minValue;
		private decimal maxValue;
		private int sortOrder;
		private string createdBy;
		private DateTime? createdDate;
		private string updatedBy;
		private DateTime? updatedDate;
		private int checkValueType;
		private string productStepCode;
		private string unit;
		private int x;
		private int y;
		private int isDisplay;
		private string condition;
		private int toolsID;
		private string frequency;
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
	
		public int KnifeStepID
		{
			get { return knifeStepID; }
			set { knifeStepID = value; }
		}
	
		public int WorkingID
		{
			get { return workingID; }
			set { workingID = value; }
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
	
		public decimal MinValue
		{
			get { return minValue; }
			set { minValue = value; }
		}
	
		public decimal MaxValue
		{
			get { return maxValue; }
			set { maxValue = value; }
		}
	
		public int SortOrder
		{
			get { return sortOrder; }
			set { sortOrder = value; }
		}
	
		public string CreatedBy
		{
			get { return createdBy; }
			set { createdBy = value; }
		}
	
		public DateTime? CreatedDate
		{
			get { return createdDate; }
			set { createdDate = value; }
		}
	
		public string UpdatedBy
		{
			get { return updatedBy; }
			set { updatedBy = value; }
		}
	
		public DateTime? UpdatedDate
		{
			get { return updatedDate; }
			set { updatedDate = value; }
		}
	
		public int CheckValueType
		{
			get { return checkValueType; }
			set { checkValueType = value; }
		}
	
		public string ProductStepCode
		{
			get { return productStepCode; }
			set { productStepCode = value; }
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
	
		public int ToolsID
		{
			get { return toolsID; }
			set { toolsID = value; }
		}
	
		public string Frequency
		{
			get { return frequency; }
			set { frequency = value; }
		}
	
	}
}
	