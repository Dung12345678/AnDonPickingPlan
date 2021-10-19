
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class UserRightDistributionBO : BaseBO
	{
		private UserRightDistributionFacade facade = UserRightDistributionFacade.Instance;
		protected static UserRightDistributionBO instance = new UserRightDistributionBO();

		protected UserRightDistributionBO()
		{
			this.baseFacade = facade;
		}

		public static UserRightDistributionBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	