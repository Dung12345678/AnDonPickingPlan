
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class WorkingBO : BaseBO
	{
		private WorkingFacade facade = WorkingFacade.Instance;
		protected static WorkingBO instance = new WorkingBO();

		protected WorkingBO()
		{
			this.baseFacade = facade;
		}

		public static WorkingBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	