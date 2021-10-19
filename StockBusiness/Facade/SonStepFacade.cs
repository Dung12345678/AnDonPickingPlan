
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class SonStepFacade : BaseFacade
	{
		protected static SonStepFacade instance = new SonStepFacade(new SonStepModel());
		protected SonStepFacade(SonStepModel model) : base(model)
		{
		}
		public static SonStepFacade Instance
		{
			get { return instance; }
		}
		protected SonStepFacade():base() 
		{ 
		} 
	
	}
}
	