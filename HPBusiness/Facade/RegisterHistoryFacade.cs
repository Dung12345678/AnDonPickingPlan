
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class RegisterHistoryFacade : BaseFacade
	{
		protected static RegisterHistoryFacade instance = new RegisterHistoryFacade(new RegisterHistoryModel());
		protected RegisterHistoryFacade(RegisterHistoryModel model) : base(model)
		{
		}
		public static RegisterHistoryFacade Instance
		{
			get { return instance; }
		}
		protected RegisterHistoryFacade():base() 
		{ 
		} 
	
	}
}
	