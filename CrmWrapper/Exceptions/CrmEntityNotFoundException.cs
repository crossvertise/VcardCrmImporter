namespace CrmWrapper.Exceptions
{
    public class CrmEntityNotFoundException : CrmIntegrationException
    {
        public CrmEntityNotFoundException(string message)
            : base(message)
        {
        }
    }
}