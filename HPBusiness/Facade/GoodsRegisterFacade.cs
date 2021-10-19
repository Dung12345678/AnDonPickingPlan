
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class GoodsRegisterFacade : BaseFacade
	{
		protected static GoodsRegisterFacade instance = new GoodsRegisterFacade(new GoodsRegisterModel());
		protected GoodsRegisterFacade(GoodsRegisterModel model) : base(model)
		{
		}
		public static GoodsRegisterFacade Instance
		{
			get { return instance; }
		}
		protected GoodsRegisterFacade():base() 
		{ 
		} 
	
	}
}
	