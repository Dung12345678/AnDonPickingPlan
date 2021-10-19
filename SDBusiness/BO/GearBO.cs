
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class GearBO : BaseBO
	{
		private GearFacade facade = GearFacade.Instance;
		protected static GearBO instance = new GearBO();

		protected GearBO()
		{
			this.baseFacade = facade;
		}

		public static GearBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	