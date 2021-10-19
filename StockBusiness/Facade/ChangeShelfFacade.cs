
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class ChangeShelfFacade : BaseFacade
	{
		protected static ChangeShelfFacade instance = new ChangeShelfFacade(new ChangeShelfModel());
		protected ChangeShelfFacade(ChangeShelfModel model) : base(model)
		{
		}
		public static ChangeShelfFacade Instance
		{
			get { return instance; }
		}
		protected ChangeShelfFacade():base() 
		{ 
		} 
	
	}
}
	