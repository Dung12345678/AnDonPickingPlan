
using System;
namespace ST.Model
{
	public class ImageRegisterModel : BaseModel
	{
		private int iD;
		private int partID;
		private int errorID;
		private string pathPatternImage;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public int PartID
		{
			get { return partID; }
			set { partID = value; }
		}
	
		public int ErrorID
		{
			get { return errorID; }
			set { errorID = value; }
		}
	
		public string PathPatternImage
		{
			get { return pathPatternImage; }
			set { pathPatternImage = value; }
		}
	
	}
}
	