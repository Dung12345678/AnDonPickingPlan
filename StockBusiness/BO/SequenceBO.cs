
using System;
using System.Collections;
using ST.Facade;
using ST.Model;
namespace ST.Business
{

	
	public class SequenceBO : BaseBO
	{
		private SequenceFacade facade = SequenceFacade.Instance;
		protected static SequenceBO instance = new SequenceBO();

		protected SequenceBO()
		{
			this.baseFacade = facade;
		}

		public static SequenceBO Instance
		{
			get { return instance; }
		}
		
	
	}
}
	