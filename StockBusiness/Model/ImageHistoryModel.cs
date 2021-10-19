
using System;
namespace ST.Model
{
	public class ImageHistoryModel : BaseModel
	{
		private int iD;
		private int partID;
		private int errorID;
		private string qrCode;
		private string path;
		private string workerCode;
		private string orderCode;
		private string qrCodePart;
		private DateTime? createDate;
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
	
		public string QrCode
		{
			get { return qrCode; }
			set { qrCode = value; }
		}
	
		public string Path
		{
			get { return path; }
			set { path = value; }
		}
	
		public string WorkerCode
		{
			get { return workerCode; }
			set { workerCode = value; }
		}
	
		public string OrderCode
		{
			get { return orderCode; }
			set { orderCode = value; }
		}
	
		public string QrCodePart
		{
			get { return qrCodePart; }
			set { qrCodePart = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
	}
}
	