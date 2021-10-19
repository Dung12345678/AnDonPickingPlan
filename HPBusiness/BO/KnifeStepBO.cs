
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class KnifeStepBO : BaseBO
	{
		private KnifeStepFacade facade = KnifeStepFacade.Instance;
		protected static KnifeStepBO instance = new KnifeStepBO();

		protected KnifeStepBO()
		{
			this.baseFacade = facade;
		}

		public static KnifeStepBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	