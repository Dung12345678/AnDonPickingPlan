
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class AssemblyStockFacade : BaseFacade
	{
		protected static AssemblyStockFacade instance = new AssemblyStockFacade(new AssemblyStockModel());
		protected AssemblyStockFacade(AssemblyStockModel model) : base(model)
		{
		}
		public static AssemblyStockFacade Instance
		{
			get { return instance; }
		}
		protected AssemblyStockFacade():base() 
		{ 
		} 
	
	}
}
	