
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ProductTypeFacade : BaseFacade
	{
		protected static ProductTypeFacade instance = new ProductTypeFacade(new ProductTypeModel());
		protected ProductTypeFacade(ProductTypeModel model) : base(model)
		{
		}
		public static ProductTypeFacade Instance
		{
			get { return instance; }
		}
		protected ProductTypeFacade():base() 
		{ 
		} 
	
	}
}
	