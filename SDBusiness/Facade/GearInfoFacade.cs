
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class GearInfoFacade : BaseFacade
	{
		protected static GearInfoFacade instance = new GearInfoFacade(new GearInfoModel());
		protected GearInfoFacade(GearInfoModel model) : base(model)
		{
		}
		public static GearInfoFacade Instance
		{
			get { return instance; }
		}
		protected GearInfoFacade():base() 
		{ 
		} 
	
	}
}
	