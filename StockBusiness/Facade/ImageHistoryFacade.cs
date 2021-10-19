
using System.Collections;
using ST.Model;
namespace ST.Facade
{
	
	public class ImageHistoryFacade : BaseFacade
	{
		protected static ImageHistoryFacade instance = new ImageHistoryFacade(new ImageHistoryModel());
		protected ImageHistoryFacade(ImageHistoryModel model) : base(model)
		{
		}
		public static ImageHistoryFacade Instance
		{
			get { return instance; }
		}
		protected ImageHistoryFacade():base() 
		{ 
		} 
	
	}
}
	