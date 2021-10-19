
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class UserGroupFacade : BaseFacade
	{
		protected static UserGroupFacade instance = new UserGroupFacade(new UserGroupModel());
		protected UserGroupFacade(UserGroupModel model) : base(model)
		{
		}
		public static UserGroupFacade Instance
		{
			get { return instance; }
		}
		protected UserGroupFacade():base() 
		{ 
		} 
	
	}
}
	