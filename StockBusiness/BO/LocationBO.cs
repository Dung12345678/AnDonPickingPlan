
using System;
using System.Collections;
using ST.Facade;
using ST.Model;
namespace ST.Business
{

	
	public class LocationBO : BaseBO
	{
		private LocationFacade facade = LocationFacade.Instance;
		protected static LocationBO instance = new LocationBO();

		protected LocationBO()
		{
			this.baseFacade = facade;
		}

		public static LocationBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	