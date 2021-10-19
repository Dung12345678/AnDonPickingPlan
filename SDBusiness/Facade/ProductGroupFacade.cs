
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ProductGroupFacade : BaseFacade
	{
		protected static ProductGroupFacade instance = new ProductGroupFacade(new ProductGroupModel());
		protected ProductGroupFacade(ProductGroupModel model) : base(model)
		{
		}
		public static ProductGroupFacade Instance
		{
			get { return instance; }
		}
		protected ProductGroupFacade():base() 
		{ 
		} 
	
	}
}
	