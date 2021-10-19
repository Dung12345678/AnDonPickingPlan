
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class ToolDetailHistoryBO : BaseBO
	{
		private ToolDetailHistoryFacade facade = ToolDetailHistoryFacade.Instance;
		protected static ToolDetailHistoryBO instance = new ToolDetailHistoryBO();

		protected ToolDetailHistoryBO()
		{
			this.baseFacade = facade;
		}

		public static ToolDetailHistoryBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	