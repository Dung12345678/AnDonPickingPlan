
using System;
namespace HP.Model
{
	public class WorkingStepModel : BaseModel
	{
		private int iD;
		private int knifeGroupID;
		private string workingStepCode;
		private string description;
		private int sortOrder;
		private DateTime? createdDate;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int KnifeGroupID
		{
			get { return knifeGroupID; }
			set { knifeGroupID = value; }
		}
	
		public string WorkingStepCode
		{
			get { return workingStepCode; }
			set { workingStepCode = value; }
		}
	
		public string Description
		{
			get { return description; }
			set { description = value; }
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
	
	}
}
	