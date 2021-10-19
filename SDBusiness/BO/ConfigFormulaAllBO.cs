
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class ConfigFormulaAllBO : BaseBO
	{
		private ConfigFormulaAllFacade facade = ConfigFormulaAllFacade.Instance;
		protected static ConfigFormulaAllBO instance = new ConfigFormulaAllBO();

		protected ConfigFormulaAllBO()
		{
			this.baseFacade = facade;
		}

		public static ConfigFormulaAllBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	