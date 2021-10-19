
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class KnifeGroupFacade : BaseFacade
	{
		protected static KnifeGroupFacade instance = new KnifeGroupFacade(new KnifeGroupModel());
		protected KnifeGroupFacade(KnifeGroupModel model) : base(model)
		{
		}
		public static KnifeGroupFacade Instance
		{
			get { return instance; }
		}
		protected KnifeGroupFacade():base() 
		{ 
		} 
	
	}
}
	