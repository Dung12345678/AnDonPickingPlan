
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class FormAndFunctionGroupBO : BaseBO
	{
		private FormAndFunctionGroupFacade facade = FormAndFunctionGroupFacade.Instance;
		protected static FormAndFunctionGroupBO instance = new FormAndFunctionGroupBO();

		protected FormAndFunctionGroupBO()
		{
			this.baseFacade = facade;
		}

		public static FormAndFunctionGroupBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	