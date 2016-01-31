namespace CrmWrapper.Exceptions
{
    public class CrmAmbiguousMatchException : CrmIntegrationException
    {
        public CrmAmbiguousMatchException(string message)
            : base(message)
        {
        }
    }
}