
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class CheckHistoryBO : BaseBO
	{
		private CheckHistoryFacade facade = CheckHistoryFacade.Instance;
		protected static CheckHistoryBO instance = new CheckHistoryBO();

		protected CheckHistoryBO()
		{
			this.baseFacade = facade;
		}

		public static CheckHistoryBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	