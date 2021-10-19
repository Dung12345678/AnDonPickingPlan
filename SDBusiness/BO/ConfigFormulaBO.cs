
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{
	
	public class ConfigFormulaBO : BaseBO
	{
		
		private ConfigFormulaFacade facade = ConfigFormulaFacade.Instance;
		protected static ConfigFormulaBO instance = new ConfigFormulaBO();

		protected ConfigFormulaBO()
		{
			this.baseFacade = facade;
		}

		public static ConfigFormulaBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	