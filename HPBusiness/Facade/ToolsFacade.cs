
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class ToolsFacade : BaseFacade
	{
		protected static ToolsFacade instance = new ToolsFacade(new ToolsModel());
		protected ToolsFacade(ToolsModel model) : base(model)
		{
		}
		public static ToolsFacade Instance
		{
			get { return instance; }
		}
		protected ToolsFacade():base() 
		{ 
		} 
	
	}
}
	