
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class RegisterBO : BaseBO
	{
		private RegisterFacade facade = RegisterFacade.Instance;
		protected static RegisterBO instance = new RegisterBO();

		protected RegisterBO()
		{
			this.baseFacade = facade;
		}

		public static RegisterBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	