
using System;
namespace ST.Model
{
	public class SequenceModel : BaseModel
	{
		private int iD;
		private string sequenceCode;
		private string sequenceName;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string SequenceCode
		{
			get { return sequenceCode; }
			set { sequenceCode = value; }
		}
	
		public string SequenceName
		{
			get { return sequenceName; }
			set { sequenceName = value; }
		}
	
	}
}
	