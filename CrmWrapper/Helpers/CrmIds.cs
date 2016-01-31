namespace CrmWrapper.Helpers
{
    using System;
    using System.Collections.Generic;

    public static class CrmIds
    {
        public static readonly Guid CompanyTypePublisher = Guid.Parse("a98e29b5-7b58-e311-9405-00155d028b03");

        public static readonly Guid CompanyTypeCustomer  = Guid.Parse("63C9CF88-7C83-E311-9407-00155D028B03");

        public static readonly Guid AccountSourceUnknown = Guid.Parse("4039718a-7078-e311-9407-00155d028b03");

        private static readonly Dictionary<string, int> SalutationLookup = new Dictionary<string, int>
            {
                { "Hr.", 772600000 },
                { "Fr.", 772600001 },
                { "Hr. Dr.", 772600002 },
                { "Fr. Dr.", 772600003 },
                { "Hr. Prof. Dr.", 772600004 },
                { "Fr. Prof. Dr.", 772600005 },
                { "Hr. Prof.", 772600006 },
                { "Fr. Prof.", 772600007 }
            };

        public static int GetSalutationId(string salutation)
        {
            int result;
            if (SalutationLookup.TryGetValue(salutation, out result))
            {
                return result;
            }

            return 0;
        }
    }
}
