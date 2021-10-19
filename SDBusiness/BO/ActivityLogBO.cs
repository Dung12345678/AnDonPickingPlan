
using System;
using System.Collections;
using SD.BO;
using SD.Exceptions;
using SD.Utils;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class ActivityLogBO : BaseBO
	{
		private ActivityLogFacade facade = ActivityLogFacade.Instance;
		protected static ActivityLogBO instance = new ActivityLogBO();

		protected ActivityLogBO()
		{
			this.baseFacade = facade;
		}

		public static ActivityLogBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	