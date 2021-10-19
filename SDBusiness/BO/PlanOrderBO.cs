
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class PlanOrderBO : BaseBO
	{
		private PlanOrderFacade facade = PlanOrderFacade.Instance;
		protected static PlanOrderBO instance = new PlanOrderBO();

		protected PlanOrderBO()
		{
			this.baseFacade = facade;
		}

		public static PlanOrderBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	