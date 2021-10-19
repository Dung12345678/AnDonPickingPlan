
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class KnifeJigStepFacade : BaseFacade
	{
		protected static KnifeJigStepFacade instance = new KnifeJigStepFacade(new KnifeJigStepModel());
		protected KnifeJigStepFacade(KnifeJigStepModel model) : base(model)
		{
		}
		public static KnifeJigStepFacade Instance
		{
			get { return instance; }
		}
		protected KnifeJigStepFacade():base() 
		{ 
		} 
	
	}
}
	