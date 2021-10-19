
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class GearWorkingFacade : BaseFacade
	{
		protected static GearWorkingFacade instance = new GearWorkingFacade(new GearWorkingModel());
		protected GearWorkingFacade(GearWorkingModel model) : base(model)
		{
		}
		public static GearWorkingFacade Instance
		{
			get { return instance; }
		}
		protected GearWorkingFacade():base() 
		{ 
		} 
	
	}
}
	