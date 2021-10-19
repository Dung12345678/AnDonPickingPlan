
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class ProductionPlanBO : BaseBO
	{
		private ProductionPlanFacade facade = ProductionPlanFacade.Instance;
		protected static ProductionPlanBO instance = new ProductionPlanBO();

		protected ProductionPlanBO()
		{
			this.baseFacade = facade;
		}

		public static ProductionPlanBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	