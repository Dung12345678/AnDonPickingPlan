
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class CustomerFacade : BaseFacade
	{
		protected static CustomerFacade instance = new CustomerFacade(new CustomerModel());
		protected CustomerFacade(CustomerModel model) : base(model)
		{
		}
		public static CustomerFacade Instance
		{
			get { return instance; }
		}
		protected CustomerFacade():base() 
		{ 
		} 
	
	}
}
	