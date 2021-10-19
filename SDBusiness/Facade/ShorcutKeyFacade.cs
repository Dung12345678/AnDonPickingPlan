
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ShorcutKeyFacade : BaseFacade
	{
		protected static ShorcutKeyFacade instance = new ShorcutKeyFacade(new ShorcutKeyModel());
		protected ShorcutKeyFacade(ShorcutKeyModel model) : base(model)
		{
		}
		public static ShorcutKeyFacade Instance
		{
			get { return instance; }
		}
		protected ShorcutKeyFacade():base() 
		{ 
		} 
	
	}
}
	