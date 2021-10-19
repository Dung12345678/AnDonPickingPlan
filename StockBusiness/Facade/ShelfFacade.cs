
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class ShelfFacade : BaseFacade
	{
		protected static ShelfFacade instance = new ShelfFacade(new ShelfModel());
		protected ShelfFacade(ShelfModel model) : base(model)
		{
		}
		public static ShelfFacade Instance
		{
			get { return instance; }
		}
		protected ShelfFacade():base() 
		{ 
		} 
	
	}
}
	