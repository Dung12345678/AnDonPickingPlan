
using System;
namespace BMS.Model
{
	public partial class SpeedReducerModel : BaseModel
	{
		public int ID {get; set;}
		
		public string StandValue {get; set;}
		
		public decimal Min {get; set;}
		
		public decimal Max {get; set;}
		
		public int Line {get; set;}
		
		public decimal TimeRunF {get; set;}
		
		public decimal TimeTeaching {get; set;}
		
		public decimal TimeGetVongQuay {get; set;}
		
		public decimal TimeMotorStop {get; set;}
		
		public DateTime? CreateDate {get; set;}
		
		public string Name {get; set;}
		
	}
}
	