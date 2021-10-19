
using System.Collections;
using SD.Model;
namespace SD.Facade
{
	
	public class SettingFFTFacade : BaseFacade
	{
		protected static SettingFFTFacade instance = new SettingFFTFacade(new SettingFFTModel());
		protected SettingFFTFacade(SettingFFTModel model) : base(model)
		{
		}
		public static SettingFFTFacade Instance
		{
			get { return instance; }
		}
		protected SettingFFTFacade():base() 
		{ 
		} 
	
	}
}
	