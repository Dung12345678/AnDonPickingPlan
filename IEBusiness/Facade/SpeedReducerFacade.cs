
using System.Collections;
using IE.Model;
namespace IE.Facade
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
	