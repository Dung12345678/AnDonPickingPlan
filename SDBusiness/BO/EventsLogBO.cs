using System;
using System.Collections;
using SD.BO;
using SD.Exceptions;
using SD.Utils;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class EventsLogBO : BaseBO
	{
		private EventsLogFacade facade = EventsLogFacade.Instance;
		protected static EventsLogBO instance = new EventsLogBO();

		protected EventsLogBO()
		{
			this.baseFacade = facade;
		}

		public static EventsLogBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	