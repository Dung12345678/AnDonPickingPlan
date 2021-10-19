
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class UserGroupRightDistributionBO : BaseBO
	{
		private UserGroupRightDistributionFacade facade = UserGroupRightDistributionFacade.Instance;
		protected static UserGroupRightDistributionBO instance = new UserGroupRightDistributionBO();

		protected UserGroupRightDistributionBO()
		{
			this.baseFacade = facade;
		}

		public static UserGroupRightDistributionBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	