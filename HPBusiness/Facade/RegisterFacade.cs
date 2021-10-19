
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class RegisterFacade : BaseFacade
	{
		protected static RegisterFacade instance = new RegisterFacade(new RegisterModel());
		protected RegisterFacade(RegisterModel model) : base(model)
		{
		}
		public static RegisterFacade Instance
		{
			get { return instance; }
		}
		protected RegisterFacade():base() 
		{ 
		} 
	
	}
}
	