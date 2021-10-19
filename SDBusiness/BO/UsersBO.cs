
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class UsersBO : BaseBO
	{
		private UsersFacade facade = UsersFacade.Instance;
		protected static UsersBO instance = new UsersBO();

		protected UsersBO()
		{
			this.baseFacade = facade;
		}

		public static UsersBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	