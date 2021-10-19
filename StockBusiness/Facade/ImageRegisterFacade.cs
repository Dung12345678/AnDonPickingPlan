
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class ImageRegisterFacade : BaseFacade
	{
		protected static ImageRegisterFacade instance = new ImageRegisterFacade(new ImageRegisterModel());
		protected ImageRegisterFacade(ImageRegisterModel model) : base(model)
		{
		}
		public static ImageRegisterFacade Instance
		{
			get { return instance; }
		}
		protected ImageRegisterFacade():base() 
		{ 
		} 
	
	}
}
	