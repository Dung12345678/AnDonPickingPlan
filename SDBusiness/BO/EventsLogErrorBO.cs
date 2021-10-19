
using System;
using System.Collections;
using SD.Exceptions;
using SD.Utils;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class EventsLogErrorBO : BaseBO
	{
		private EventsLogErrorFacade facade = EventsLogErrorFacade.Instance;
		protected static EventsLogErrorBO instance = new EventsLogErrorBO();

		protected EventsLogErrorBO()
		{
			this.baseFacade = facade;
		}

		public static EventsLogErrorBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	