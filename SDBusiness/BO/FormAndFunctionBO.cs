
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class FormAndFunctionBO : BaseBO
	{
		private FormAndFunctionFacade facade = FormAndFunctionFacade.Instance;
		protected static FormAndFunctionBO instance = new FormAndFunctionBO();

		protected FormAndFunctionBO()
		{
			this.baseFacade = facade;
		}

		public static FormAndFunctionBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	