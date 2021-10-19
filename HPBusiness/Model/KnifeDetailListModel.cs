
using System;
namespace HP.Model
{
	public class KnifeDetailListModel : BaseModel
	{
		private int iD;
		private int knifeGroupID;
		private string knifeCode;
		private string knifeName;
		private string description;
		private DateTime? importedDate;
		private int workerID;
		private int sTD;
		private int aTC;
		private bool status;
		private int currentSTD;
		private int currentATC;
		private DateTime? updatedDate;
		private int type;
		private int parentID;
		private int qty;
		private int totalNumber;
		private int remainQty;
		private string frequencyCheck;
		private decimal valueFrequency;
		private int typeFrequencyCheck;
		private DateTime? dateOld;
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
	
		public string KnifeCode
		{
			get { return knifeCode; }
			set { knifeCode = value; }
		}
	
		public string KnifeName
		{
			get { return knifeName; }
			set { knifeName = value; }
		}
	
		public string Description
		{
			get { return description; }
			set { description = value; }
		}
	
		public DateTime? ImportedDate
		{
			get { return importedDate; }
			set { importedDate = value; }
		}
	
		public int WorkerID
		{
			get { return workerID; }
			set { workerID = value; }
		}
	
		public int STD
		{
			get { return sTD; }
			set { sTD = value; }
		}
	
		public int ATC
		{
			get { return aTC; }
			set { aTC = value; }
		}
	
		public bool Status
		{
			get { return status; }
			set { status = value; }
		}
	
		public int CurrentSTD
		{
			get { return currentSTD; }
			set { currentSTD = value; }
		}
	
		public int CurrentATC
		{
			get { return currentATC; }
			set { currentATC = value; }
		}
	
		public DateTime? UpdatedDate
		{
			get { return updatedDate; }
			set { updatedDate = value; }
		}
	
		public int Type
		{
			get { return type; }
			set { type = value; }
		}
	
		public int ParentID
		{
			get { return parentID; }
			set { parentID = value; }
		}
	
		public int Qty
		{
			get { return qty; }
			set { qty = value; }
		}
	
		public int TotalNumber
		{
			get { return totalNumber; }
			set { totalNumber = value; }
		}
	
		public int RemainQty
		{
			get { return remainQty; }
			set { remainQty = value; }
		}
	
		public string FrequencyCheck
		{
			get { return frequencyCheck; }
			set { frequencyCheck = value; }
		}
	
		public decimal ValueFrequency
		{
			get { return valueFrequency; }
			set { valueFrequency = value; }
		}
	
		public int TypeFrequencyCheck
		{
			get { return typeFrequencyCheck; }
			set { typeFrequencyCheck = value; }
		}
	
		public DateTime? DateOld
		{
			get { return dateOld; }
			set { dateOld = value; }
		}
	
	}
}
	