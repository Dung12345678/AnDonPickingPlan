
using System;
using System.Collections;
using ST.Facade;
using ST.Model;
namespace ST.Business
{

	
	public class SonStepBO : BaseBO
	{
		private SonStepFacade facade = SonStepFacade.Instance;
		protected static SonStepBO instance = new SonStepBO();

		protected SonStepBO()
		{
			this.baseFacade = facade;
		}

		public static SonStepBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	