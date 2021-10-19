
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class LocationFacade : BaseFacade
	{
		protected static LocationFacade instance = new LocationFacade(new LocationModel());
		protected LocationFacade(LocationModel model) : base(model)
		{
		}
		public static LocationFacade Instance
		{
			get { return instance; }
		}
		protected LocationFacade():base() 
		{ 
		} 
	
	}
}
	