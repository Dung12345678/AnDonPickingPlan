
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class KnifeJigStepBO : BaseBO
	{
		private KnifeJigStepFacade facade = KnifeJigStepFacade.Instance;
		protected static KnifeJigStepBO instance = new KnifeJigStepBO();

		protected KnifeJigStepBO()
		{
			this.baseFacade = facade;
		}

		public static KnifeJigStepBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	