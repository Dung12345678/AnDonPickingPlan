
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class SupplierBO : BaseBO
	{
		private SupplierFacade facade = SupplierFacade.Instance;
		protected static SupplierBO instance = new SupplierBO();

		protected SupplierBO()
		{
			this.baseFacade = facade;
		}

		public static SupplierBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	