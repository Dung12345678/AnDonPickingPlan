
using System;
namespace HP.Model
{
	public class ToolsModel : BaseModel
	{
		private int iD;
		private string code;
		private string name;
		private double std;
		private decimal min;
		private decimal max;
		private int type;
		private string nameDisplay;
		private string conditionDisplay;
		private DateTime? createDate;
		private int groupID;
		private DateTime? dateOld;
		private string frequencyCheck;
		private int typeFrequencyCheck;
		private decimal valueFrequency;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string Code
		{
			get { return code; }
			set { code = value; }
		}
	
		public string Name
		{
			get { return name; }
			set { name = value; }
		}
	
		public double Std
		{
			get { return std; }
			set { std = value; }
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
	
		public int Type
		{
			get { return type; }
			set { type = value; }
		}
	
		public string NameDisplay
		{
			get { return nameDisplay; }
			set { nameDisplay = value; }
		}
	
		public string ConditionDisplay
		{
			get { return conditionDisplay; }
			set { conditionDisplay = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public int GroupID
		{
			get { return groupID; }
			set { groupID = value; }
		}
	
		public DateTime? DateOld
		{
			get { return dateOld; }
			set { dateOld = value; }
		}
	
		public string FrequencyCheck
		{
			get { return frequencyCheck; }
			set { frequencyCheck = value; }
		}
	
		public int TypeFrequencyCheck
		{
			get { return typeFrequencyCheck; }
			set { typeFrequencyCheck = value; }
		}
	
		public decimal ValueFrequency
		{
			get { return valueFrequency; }
			set { valueFrequency = value; }
		}
	
	}
}
	