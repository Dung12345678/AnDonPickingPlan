
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class SkipCDFacade : BaseFacade
	{
		protected static SkipCDFacade instance = new SkipCDFacade(new SkipCDModel());
		protected SkipCDFacade(SkipCDModel model) : base(model)
		{
		}
		public static SkipCDFacade Instance
		{
			get { return instance; }
		}
		protected SkipCDFacade():base() 
		{ 
		} 
	
	}
}
	