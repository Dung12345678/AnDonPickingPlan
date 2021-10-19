
using System;
namespace HP.Model
{
	public class KnifeJigStepModel : BaseModel
	{
		private int iD;
		private int workingStepID;
		private int knifeDetailListID;
		private string productStepCode;
		private string description;
		private int parentID;
		private int sortOrder;
		private DateTime? createdDate;
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
	
		public int KnifeDetailListID
		{
			get { return knifeDetailListID; }
			set { knifeDetailListID = value; }
		}
	
		public string ProductStepCode
		{
			get { return productStepCode; }
			set { productStepCode = value; }
		}
	
		public string Description
		{
			get { return description; }
			set { description = value; }
		}
	
		public int ParentID
		{
			get { return parentID; }
			set { parentID = value; }
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
	