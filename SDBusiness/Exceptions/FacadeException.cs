using System;
using SD.Utils;
using SD.Facade;
using SD.Exceptions;
using SD.Business;
using SD.Model;
namespace SD.Exceptions
{
	public class FacadeException : Exception
	{
		public FacadeException(String message) : base(message)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public FacadeException(Exception e)
		{			
		}
	}
}