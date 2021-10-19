
using System;
using System.Collections;
using ST.Facade;
using ST.Model;
namespace ST.Business
{

	
	public class ImageHistoryBO : BaseBO
	{
		private ImageHistoryFacade facade = ImageHistoryFacade.Instance;
		protected static ImageHistoryBO instance = new ImageHistoryBO();

		protected ImageHistoryBO()
		{
			this.baseFacade = facade;
		}

		public static ImageHistoryBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	