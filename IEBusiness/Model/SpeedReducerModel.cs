
using System;
namespace IE.Model
{
	public class SpeedReducerModel : BaseModel
	{
		private int iD;
		private string standValue;
		private decimal min;
		private decimal max;
		private int line;
		private decimal timeRunF;
		private decimal timeRunR;
		private decimal timeTeaching;
		private decimal timeGetVongQuay;
		private decimal timeMotorStop;
		private DateTime? createDate;
		private string name;
		public int ID
		{
			get { return iD; }
			set { iD = value; }
		}
	
		public string StandValue
		{
			get { return standValue; }
			set { standValue = value; }
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
	
		public int Line
		{
			get { return line; }
			set { line = value; }
		}
	
		public decimal TimeRunF
		{
			get { return timeRunF; }
			set { timeRunF = value; }
		}
	
		public decimal TimeRunR
		{
			get { return timeRunR; }
			set { timeRunR = value; }
		}
	
		public decimal TimeTeaching
		{
			get { return timeTeaching; }
			set { timeTeaching = value; }
		}
	
		public decimal TimeGetVongQuay
		{
			get { return timeGetVongQuay; }
			set { timeGetVongQuay = value; }
		}
	
		public decimal TimeMotorStop
		{
			get { return timeMotorStop; }
			set { timeMotorStop = value; }
		}
	
		public DateTime? CreateDate
		{
			get { return createDate; }
			set { createDate = value; }
		}
	
		public string Name
		{
			get { return name; }
			set { name = value; }
		}
	
	}
}
	