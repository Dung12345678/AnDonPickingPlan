
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class CheckHistoryFacade : BaseFacade
	{
		protected static CheckHistoryFacade instance = new CheckHistoryFacade(new CheckHistoryModel());
		protected CheckHistoryFacade(CheckHistoryModel model) : base(model)
		{
		}
		public static CheckHistoryFacade Instance
		{
			get { return instance; }
		}
		protected CheckHistoryFacade():base() 
		{ 
		} 
	
	}
}
	