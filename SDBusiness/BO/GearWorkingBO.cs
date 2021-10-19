
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class GearWorkingBO : BaseBO
	{
		private GearWorkingFacade facade = GearWorkingFacade.Instance;
		protected static GearWorkingBO instance = new GearWorkingBO();

		protected GearWorkingBO()
		{
			this.baseFacade = facade;
		}

		public static GearWorkingBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	