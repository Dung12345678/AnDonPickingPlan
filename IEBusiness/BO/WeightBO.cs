
using System;
using System.Collections;
using IE.Facade;
using IE.Model;
namespace IE.Business
{
	
	public class WeightBO : BaseBO
	{
		private WeightFacade facade = WeightFacade.Instance;
		protected static WeightBO instance = new WeightBO();

		protected WeightBO()
		{
			this.baseFacade = facade;
		}

		public static WeightBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	