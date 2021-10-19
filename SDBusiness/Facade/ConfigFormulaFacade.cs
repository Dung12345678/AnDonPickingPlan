
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class ConfigFormulaFacade : BaseFacade
	{
		protected static ConfigFormulaFacade instance = new ConfigFormulaFacade(new ConfigFormulaModel());
		protected ConfigFormulaFacade(ConfigFormulaModel model) : base(model)
		{
		}
		public static ConfigFormulaFacade Instance
		{
			get { return instance; }
		}
		protected ConfigFormulaFacade():base() 
		{ 
		} 
	
	}
}
	