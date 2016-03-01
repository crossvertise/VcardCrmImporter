namespace VcardCrmImporter.Services
{
    using System;
    using System.Linq;
    using System.Reflection;

    using CrmOrganizationClasses;

    // These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
    // located in the SDK\bin folder of the SDK download.
    using CrmWrapper.Extensions;
    using CrmWrapper.Helpers;

    using Microsoft.Azure;
    using Microsoft.Xrm.Client;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Client;

    using Thought.vCards;

    public class CrmVcardUpdateService : IDisposable
    {
        private readonly OrganizationServiceProxy service;

        private readonly OrganizationServiceContext orgContext;

        private readonly SystemUser impersonatedUser;

        public CrmVcardUpdateService(string userEmail)
        {
            if (string.IsNullOrWhiteSpace(userEmail))
            {
                throw new ArgumentException("userEmail must be a valid email address", "userEmail");
            }

            // Establish CRM connection
            var crmConnection = CrmConnection.Parse(CloudConfigurationManager.GetSetting("CrmConnectionString"));
            var serviceUri = new Uri(crmConnection.ServiceUri + "/XRMServices/2011/Organization.svc");
            this.service = new OrganizationServiceProxy(serviceUri, null, crmConnection.ClientCredentials, null);

            // This statement is required to enable early-bound type support.
            this.service.EnableProxyTypes(Assembly.GetAssembly(typeof(SystemUser)));

            // Create context
            this.orgContext = new OrganizationServiceContext(this.service);

            // Retrieve the system user ID of the user to impersonate.
            this.impersonatedUser = (from user in this.orgContext.CreateQuery<SystemUser>() where user.InternalEMailAddress == userEmail select user).FirstOrDefault();

            // We impersonate the user that has sent the email
            if (this.impersonatedUser == null)
            {
                throw new Exception("User not found in CRM");
            }

            this.service.CallerId = this.impersonatedUser.Id;
        }

        public string UpdateContactWithVcard(vCard vcard, string filename)
        {
            // pre-checks
            if (string.IsNullOrWhiteSpace(vcard.FamilyName))
            {
                return string.Format("{0}: not imported - family name not provided", filename);
            }

            if (string.IsNullOrWhiteSpace(vcard.GivenName))
            {
                return string.Format("{0}: not imported - given name not provided", filename);
            }

            if (string.IsNullOrWhiteSpace(vcard.Organization))
            {
                return string.Format("{0}: not imported - company not provided", filename);
            }

            if (!vcard.EmailAddresses.Any())
            {
                return string.Format("{0}: not imported - email address not provided", filename);
            }

            string result;

            // Search for the company (account) in the CRM
            var companyNameShortend = CompanyNameHelper.RemoveCommonCompanySuffixes(vcard.Organization);
            var account = this.orgContext.CreateQuery<Account>().FirstOrDefault(a => a.Name.StartsWith(companyNameShortend));

            if (account != null)
            {
                // If it exists, update it
                this.UpdateAccountWithVcard(account, vcard);
                this.orgContext.UpdateObject(account);
                result = "Existing account updated";
            }
            else
            {
                // If it not yet exists, create it
                account = this.UpdateAccountWithVcard(null, vcard);
                account.OwnerId = new EntityReference(this.impersonatedUser.LogicalName, this.impersonatedUser.Id);
                this.orgContext.AddObject(account);
                result = "New account created";
            }

            // Save the account, so it receices an ID
            this.orgContext.SaveChanges();

            Contact contact = null;

            // Try to get the contact by Firstname, Lastname and Company name
            if (vcard.GivenName != null && vcard.FamilyName != null && vcard.Organization != null)
            {
                contact = (from c in this.orgContext.CreateQuery<Contact>()
                            where
                                c.FirstName == vcard.GivenName.Trim() && c.LastName == vcard.FamilyName.Trim()
                                && c.ParentCustomerId == new EntityReference(account.LogicalName, account.Id)
                            select c).FirstOrDefault();
            }

            // Try to search contact by email
            if (vcard.EmailAddresses.Any() && vcard.EmailAddresses[0].Address != null)
            {
                contact = contact ?? this.orgContext.CreateQuery<Contact>().FirstOrDefault(c => c.EMailAddress1 == vcard.EmailAddresses[0].Address.Trim());
            }

            if (contact != null)
            {
                // If the contact is found, update it
                this.UpdateContactWithVcard(contact, vcard);
                contact.ParentCustomerId = new EntityReference(account.LogicalName, account.Id);
                this.orgContext.UpdateObject(contact);
                result += ", Existing contact updated";
            }
            else
            {
                // If the contact was not found, insert it
                contact = this.UpdateContactWithVcard(null, vcard);
                contact.ParentCustomerId = new EntityReference(account.LogicalName, account.Id);
                contact.OwnerId = new EntityReference(this.impersonatedUser.LogicalName, this.impersonatedUser.Id);
                this.orgContext.AddObject(contact);
                result += ", New contact created";
            }

            // Save the contact
            this.orgContext.SaveChanges();

            return filename + ": " + result;
        }

        public void Dispose()
        {
            this.orgContext.Dispose();
            this.service.Dispose();
        }

        private Contact UpdateContactWithVcard(Contact contact, vCard vcard)
        {
            if (contact == null)
            {
                contact = new Contact();
            }

            contact.FirstName = vcard.GivenName.Trim();
            contact.LastName = vcard.FamilyName.Trim();
            
            contact.Telephone1 = vcard.Phones.Where(p => p.IsWork).Select(p => p.FullNumber).FirstOrDefault();
            contact.Fax = vcard.Phones.Where(p => p.IsFax).Select(p => p.FullNumber).FirstOrDefault();
            contact.MobilePhone = vcard.Phones.Where(p => p.IsCellular).Select(p => p.FullNumber).FirstOrDefault();
            contact.EMailAddress1 = vcard.EmailAddresses.Select(e => e.Address).FirstOrDefault();
            contact.WebSiteUrl = vcard.Websites.Select(w => w.Url).FirstOrDefault();

            contact.Address1_Line1 = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Street).FirstOrDefault();
            contact.Address1_City = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.City).FirstOrDefault();
            contact.Address1_StateOrProvince = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Region).FirstOrDefault();
            contact.Address1_Country = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Country).FirstOrDefault();
            contact.Address1_PostalCode = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.PostalCode).FirstOrDefault();
            contact.xv_Land = contact.Address1_Country != null ? this.service.GetCountryReferenceByName(contact.Address1_Country) : null;

            contact.Department = vcard.Department;
            contact.JobTitle = vcard.Role;
            contact.BirthDate = vcard.BirthDate;
            contact.xv_Salutation = vcard.Gender == vCardGender.Male ? new OptionSetValue(772600000) : (vcard.Gender == vCardGender.Female ? new OptionSetValue(772600001) : null);
            contact.Description += string.Join(" ", vcard.Notes.Select(n => n.Text));

            return contact;
        }

        private Account UpdateAccountWithVcard(Account account, vCard vcard)
        {
            if (account == null)
            {
                account = new Account();
            }

            account.Name = vcard.Organization;
            account.WebSiteURL = vcard.Websites.Select(w => w.Url).FirstOrDefault();

            account.Address1_Line1 = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Street).FirstOrDefault();
            account.Address1_City = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.City).FirstOrDefault();
            account.Address1_StateOrProvince = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Region).FirstOrDefault();
            account.Address1_Country = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.Country).FirstOrDefault();
            account.Address1_PostalCode = vcard.DeliveryAddresses.Where(a => a.IsWork).Select(a => a.PostalCode).FirstOrDefault();
            account.xv_Land = account.Address1_Country != null ? this.service.GetCountryReferenceByName(account.Address1_Country) : null;

            account.xv_Firmenklassifizierung = vcard.Categories.Count > 0 ? this.service.GetCompanyClassificationReferenceByName(vcard.Categories[0]) : null;

            return account;
        }
    }
}