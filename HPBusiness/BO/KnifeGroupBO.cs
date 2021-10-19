
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class KnifeGroupBO : BaseBO
	{
		private KnifeGroupFacade facade = KnifeGroupFacade.Instance;
		protected static KnifeGroupBO instance = new KnifeGroupBO();

		protected KnifeGroupBO()
		{
			this.baseFacade = facade;
		}

		public static KnifeGroupBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	