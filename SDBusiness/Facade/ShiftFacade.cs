
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ShiftFacade : BaseFacade
	{
		protected static ShiftFacade instance = new ShiftFacade(new ShiftModel());
		protected ShiftFacade(ShiftModel model) : base(model)
		{
		}
		public static ShiftFacade Instance
		{
			get { return instance; }
		}
		protected ShiftFacade():base() 
		{ 
		} 
	
	}
}
	