
using System;
using System.Collections;
using ST.Facade;
using ST.Model;
namespace ST.Business
{

	
	public class ChangeShelfBO : BaseBO
	{
		private ChangeShelfFacade facade = ChangeShelfFacade.Instance;
		protected static ChangeShelfBO instance = new ChangeShelfBO();

		protected ChangeShelfBO()
		{
			this.baseFacade = facade;
		}

		public static ChangeShelfBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	