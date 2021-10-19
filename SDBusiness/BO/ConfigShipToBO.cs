
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class ConfigShipToBO : BaseBO
	{
		private ConfigShipToFacade facade = ConfigShipToFacade.Instance;
		protected static ConfigShipToBO instance = new ConfigShipToBO();

		protected ConfigShipToBO()
		{
			this.baseFacade = facade;
		}

		public static ConfigShipToBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	