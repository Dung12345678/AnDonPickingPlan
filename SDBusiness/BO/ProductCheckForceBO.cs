
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class ProductCheckForceBO : BaseBO
	{
		private ProductCheckForceFacade facade = ProductCheckForceFacade.Instance;
		protected static ProductCheckForceBO instance = new ProductCheckForceBO();

		protected ProductCheckForceBO()
		{
			this.baseFacade = facade;
		}

		public static ProductCheckForceBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	