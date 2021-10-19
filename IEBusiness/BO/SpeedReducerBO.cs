
using System;
using System.Collections;
using IE.Facade;
using IE.Model;
namespace IE.Business
{
	
	public class SpeedReducerBO : BaseBO
	{
		private SpeedReducerFacade facade = SpeedReducerFacade.Instance;
		protected static SpeedReducerBO instance = new SpeedReducerBO();

		protected SpeedReducerBO()
		{
			this.baseFacade = facade;
		}

		public static SpeedReducerBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	