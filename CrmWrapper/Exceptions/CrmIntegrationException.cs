namespace CrmWrapper.Exceptions
{
    using System;

    public class CrmIntegrationException : Exception
    {
        public CrmIntegrationException(string message)
            : base(message)
        {
        }
    }
}