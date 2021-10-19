
using System.Collections;
using BMS.Model;
namespace BMS.Facade
{
	
	public class SpeedReducerFacade : BaseFacade
	{
		protected static SpeedReducerFacade instance = new SpeedReducerFacade(new SpeedReducerModel());
		protected SpeedReducerFacade(SpeedReducerModel model) : base(model)
		{
		}
		public static SpeedReducerFacade Instance
		{
			get { return instance; }
		}
		protected SpeedReducerFacade():base() 
		{ 
		} 
	
	}
}
	