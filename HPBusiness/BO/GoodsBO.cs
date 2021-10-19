
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class GoodsBO : BaseBO
	{
		private GoodsFacade facade = GoodsFacade.Instance;
		protected static GoodsBO instance = new GoodsBO();

		protected GoodsBO()
		{
			this.baseFacade = facade;
		}

		public static GoodsBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	