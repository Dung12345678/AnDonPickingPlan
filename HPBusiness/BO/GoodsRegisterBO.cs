
using System;
using System.Collections;
using HP.Facade;
using HP.Model;
namespace HP.Business
{
	
	public class GoodsRegisterBO : BaseBO
	{
		private GoodsRegisterFacade facade = GoodsRegisterFacade.Instance;
		protected static GoodsRegisterBO instance = new GoodsRegisterBO();

		protected GoodsRegisterBO()
		{
			this.baseFacade = facade;
		}

		public static GoodsRegisterBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	