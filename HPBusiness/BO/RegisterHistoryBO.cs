
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class RegisterHistoryBO : BaseBO
	{
		private RegisterHistoryFacade facade = RegisterHistoryFacade.Instance;
		protected static RegisterHistoryBO instance = new RegisterHistoryBO();

		protected RegisterHistoryBO()
		{
			this.baseFacade = facade;
		}

		public static RegisterHistoryBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	