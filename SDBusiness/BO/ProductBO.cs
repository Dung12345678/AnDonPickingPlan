
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class ProductBO : BaseBO
	{
		private ProductFacade facade = ProductFacade.Instance;
		protected static ProductBO instance = new ProductBO();

		protected ProductBO()
		{
			this.baseFacade = facade;
		}

		public static ProductBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	