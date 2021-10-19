
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class ProductWorkingAuditBO : BaseBO
	{
		private ProductWorkingAuditFacade facade = ProductWorkingAuditFacade.Instance;
		protected static ProductWorkingAuditBO instance = new ProductWorkingAuditBO();

		protected ProductWorkingAuditBO()
		{
			this.baseFacade = facade;
		}

		public static ProductWorkingAuditBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	