
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class KnifeStepFacade : BaseFacade
	{
		protected static KnifeStepFacade instance = new KnifeStepFacade(new KnifeStepModel());
		protected KnifeStepFacade(KnifeStepModel model) : base(model)
		{
		}
		public static KnifeStepFacade Instance
		{
			get { return instance; }
		}
		protected KnifeStepFacade():base() 
		{ 
		} 
	
	}
}
	