
using System.Collections;
using HP.Model;
namespace HP.Facade
{
	
	public class ConfigFormulaAllFacade : BaseFacade
	{
		protected static ConfigFormulaAllFacade instance = new ConfigFormulaAllFacade(new ConfigFormulaAllModel());
		protected ConfigFormulaAllFacade(ConfigFormulaAllModel model) : base(model)
		{
		}
		public static ConfigFormulaAllFacade Instance
		{
			get { return instance; }
		}
		protected ConfigFormulaAllFacade():base() 
		{ 
		} 
	
	}
}
	