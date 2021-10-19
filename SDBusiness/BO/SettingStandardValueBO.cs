
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class SettingStandardValueBO : BaseBO
	{
		private SettingStandardValueFacade facade = SettingStandardValueFacade.Instance;
		protected static SettingStandardValueBO instance = new SettingStandardValueBO();

		protected SettingStandardValueBO()
		{
			this.baseFacade = facade;
		}

		public static SettingStandardValueBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	