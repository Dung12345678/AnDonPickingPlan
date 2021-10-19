
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class ErrorFacade : BaseFacade
	{
		protected static ErrorFacade instance = new ErrorFacade(new ErrorModel());
		protected ErrorFacade(ErrorModel model) : base(model)
		{
		}
		public static ErrorFacade Instance
		{
			get { return instance; }
		}
		protected ErrorFacade():base() 
		{ 
		} 
	
	}
}
	