
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class CheckHistoryDetailBO : BaseBO
	{
		private CheckHistoryDetailFacade facade = CheckHistoryDetailFacade.Instance;
		protected static CheckHistoryDetailBO instance = new CheckHistoryDetailBO();

		protected CheckHistoryDetailBO()
		{
			this.baseFacade = facade;
		}

		public static CheckHistoryDetailBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	