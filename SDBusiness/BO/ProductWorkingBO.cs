
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class ProductWorkingBO : BaseBO
	{
		private ProductWorkingFacade facade = ProductWorkingFacade.Instance;
		protected static ProductWorkingBO instance = new ProductWorkingBO();

		protected ProductWorkingBO()
		{
			this.baseFacade = facade;
		}

		public static ProductWorkingBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	