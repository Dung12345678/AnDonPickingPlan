
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ProductStepFacade : BaseFacade
	{
		protected static ProductStepFacade instance = new ProductStepFacade(new ProductStepModel());
		protected ProductStepFacade(ProductStepModel model) : base(model)
		{
		}
		public static ProductStepFacade Instance
		{
			get { return instance; }
		}
		protected ProductStepFacade():base() 
		{ 
		} 
	
	}
}
	