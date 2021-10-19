
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class GearInfoBO : BaseBO
	{
		private GearInfoFacade facade = GearInfoFacade.Instance;
		protected static GearInfoBO instance = new GearInfoBO();

		protected GearInfoBO()
		{
			this.baseFacade = facade;
		}

		public static GearInfoBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	