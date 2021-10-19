
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class ToolsBO : BaseBO
	{
		private ToolsFacade facade = ToolsFacade.Instance;
		protected static ToolsBO instance = new ToolsBO();

		protected ToolsBO()
		{
			this.baseFacade = facade;
		}

		public static ToolsBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	