
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class GoodsFacade : BaseFacade
	{
		protected static GoodsFacade instance = new GoodsFacade(new GoodsModel());
		protected GoodsFacade(GoodsModel model) : base(model)
		{
		}
		public static GoodsFacade Instance
		{
			get { return instance; }
		}
		protected GoodsFacade():base() 
		{ 
		} 
	
	}
}
	