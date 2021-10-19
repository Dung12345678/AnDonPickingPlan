
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class ToolDetailHistoryFacade : BaseFacade
	{
		protected static ToolDetailHistoryFacade instance = new ToolDetailHistoryFacade(new ToolDetailHistoryModel());
		protected ToolDetailHistoryFacade(ToolDetailHistoryModel model) : base(model)
		{
		}
		public static ToolDetailHistoryFacade Instance
		{
			get { return instance; }
		}
		protected ToolDetailHistoryFacade():base() 
		{ 
		} 
	
	}
}
	