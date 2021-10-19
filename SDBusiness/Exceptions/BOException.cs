using System;

namespace SD.Exceptions
{
	public class BOException : Exception
	{
		public BOException(String message) : base(message)
		{
		}
		public BOException(String message, Exception e, string className)
		{
		}
	}
}