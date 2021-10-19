
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class UsersFacade : BaseFacade
	{
		protected static UsersFacade instance = new UsersFacade(new UsersModel());
		protected UsersFacade(UsersModel model) : base(model)
		{
		}
		public static UsersFacade Instance
		{
			get { return instance; }
		}
		protected UsersFacade():base() 
		{ 
		} 
	
	}
}
	