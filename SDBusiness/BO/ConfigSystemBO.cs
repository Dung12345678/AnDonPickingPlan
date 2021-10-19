
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class ConfigSystemBO : BaseBO
	{
		private ConfigSystemFacade facade = ConfigSystemFacade.Instance;
		protected static ConfigSystemBO instance = new ConfigSystemBO();

		protected ConfigSystemBO()
		{
			this.baseFacade = facade;
		}

		public static ConfigSystemBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	