
using System;
using System.Collections;
using SD.Facade;
using SD.Model;
namespace SD.Business
{

	
	public class TesttaBO : BaseBO
	{
		private TesttaFacade facade = TesttaFacade.Instance;
		protected static TesttaBO instance = new TesttaBO();

		protected TesttaBO()
		{
			this.baseFacade = facade;
		}

		public static TesttaBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	