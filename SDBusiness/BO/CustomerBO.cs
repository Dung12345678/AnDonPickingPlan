
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class CustomerBO : BaseBO
	{
		private CustomerFacade facade = CustomerFacade.Instance;
		protected static CustomerBO instance = new CustomerBO();

		protected CustomerBO()
		{
			this.baseFacade = facade;
		}

		public static CustomerBO Instance
		{
			get { return instance; }
		}
		
	    
	}
}
	