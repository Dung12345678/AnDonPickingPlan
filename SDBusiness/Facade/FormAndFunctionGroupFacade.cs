
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class FormAndFunctionGroupFacade : BaseFacade
	{
		protected static FormAndFunctionGroupFacade instance = new FormAndFunctionGroupFacade(new FormAndFunctionGroupModel());
		protected FormAndFunctionGroupFacade(FormAndFunctionGroupModel model) : base(model)
		{
		}
		public static FormAndFunctionGroupFacade Instance
		{
			get { return instance; }
		}
		protected FormAndFunctionGroupFacade():base() 
		{ 
		} 
	
	}
}
	