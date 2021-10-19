
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class CheckHistoryDetailFacade : BaseFacade
	{
		protected static CheckHistoryDetailFacade instance = new CheckHistoryDetailFacade(new CheckHistoryDetailModel());
		protected CheckHistoryDetailFacade(CheckHistoryDetailModel model) : base(model)
		{
		}
		public static CheckHistoryDetailFacade Instance
		{
			get { return instance; }
		}
		protected CheckHistoryDetailFacade():base() 
		{ 
		} 
	
	}
}
	