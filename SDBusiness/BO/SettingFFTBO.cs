
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class SettingFFTBO : BaseBO
	{
		private SettingFFTFacade facade = SettingFFTFacade.Instance;
		protected static SettingFFTBO instance = new SettingFFTBO();

		protected SettingFFTBO()
		{
			this.baseFacade = facade;
		}

		public static SettingFFTBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	