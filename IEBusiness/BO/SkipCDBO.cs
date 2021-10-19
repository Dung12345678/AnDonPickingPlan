
using System;
using System.Collections;
using IE.Facade;
using IE.Model;
namespace IE.Business
{
	
	public class SkipCDBO : BaseBO
	{
		private SkipCDFacade facade = SkipCDFacade.Instance;
		protected static SkipCDBO instance = new SkipCDBO();

		protected SkipCDBO()
		{
			this.baseFacade = facade;
		}

		public static SkipCDBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	