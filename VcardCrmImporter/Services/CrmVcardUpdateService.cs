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

    public class CrmVcardUpdateService
    {
        public async System.Threading.Tasks.Task UpdateContactWithVcard(vCard vcard, string userEmail)
        {
            //Todo: add pre-checks

            var crmConnection = CrmConnection.Parse(CloudConfigurationManager.GetSetting("CrmConnectionString"));
            var serviceUri = new Uri(crmConnection.ServiceUri + "/XRMServices/2011/Organization.svc");

            using (var service = new OrganizationServiceProxy(serviceUri, null, crmConnection.ClientCredentials, null))
            {
                // This statement is required to enable early-bound type support.
                service.EnableProxyTypes(Assembly.GetAssembly(typeof(SystemUser)));


                var orgContext = new OrganizationServiceContext(service);

                // Retrieve the system user ID of the user to impersonate.
                var impersonatedUser = (from user in orgContext.CreateQuery<SystemUser>()
                                          where user.InternalEMailAddress == userEmail 
                                          select user).FirstOrDefault();

                // We impersonate the user that has sent the email
                if (impersonatedUser == null)
                {
                    throw new Exception("User not found in CRM");
                }

                service.CallerId = impersonatedUser.Id;

                // Search for the company (account) in the CRM
                var companyNameShortend = CompanyNameHelper.RemoveCommonCompanySuffixes(vcard.Organization);
                var account = orgContext.CreateQuery<Account>().FirstOrDefault(a => a.Name.StartsWith(companyNameShortend));

                if (account != null)
                {
                    // If it exists, update it
                    this.UpdateAccountWithVcard(service, account, vcard);
                    orgContext.UpdateObject(account);
                }
                else
                {
                    // If it not yet exists, create it
                    account = this.UpdateAccountWithVcard(service, null, vcard);
                    account.OwnerId = new EntityReference(impersonatedUser.LogicalName, impersonatedUser.Id);
                    orgContext.AddObject(account);
                }

                // Save the account, so it receices an ID
                orgContext.SaveChanges();

                Contact contact = null;

                // Try to get the contact by Firstname, Lastname and Company name
                if (vcard.GivenName != null && vcard.FamilyName != null && vcard.Organization != null)
                {
                    contact = (from c in orgContext.CreateQuery<Contact>()
                               where
                                   c.FirstName == vcard.GivenName.Trim() && c.LastName == vcard.FamilyName.Trim()
                                   && c.ParentCustomerId == new EntityReference(account.LogicalName, account.Id)
                               select c).FirstOrDefault();
                }

                // Try to search contact by email
                if (vcard.EmailAddresses.Any() && vcard.EmailAddresses[0].Address != null)
                {
                    contact = contact ?? orgContext.CreateQuery<Contact>().FirstOrDefault(c => c.EMailAddress1 == vcard.EmailAddresses[0].Address.Trim());
                }

                if (contact != null)
                {
                    // If the contact is found, update it
                    this.UpdateContactWithVcard(service, contact, vcard);
                    contact.ParentCustomerId = new EntityReference(account.LogicalName, account.Id);
                    orgContext.UpdateObject(contact);
                }
                else
                {
                    // If the contact was not found, insert it
                    contact = this.UpdateContactWithVcard(service, null, vcard);
                    contact.ParentCustomerId = new EntityReference(account.LogicalName, account.Id);
                    contact.OwnerId = new EntityReference(impersonatedUser.LogicalName, impersonatedUser.Id);
                    orgContext.AddObject(contact);
                }

                // Save the contact
                orgContext.SaveChanges();
            }
        }

        private Contact UpdateContactWithVcard(IOrganizationService service, Contact contact, vCard vcard)
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
            contact.xv_Land = contact.Address1_Country != null ? service.GetCountryReferenceByName(contact.Address1_Country) : null;

            contact.Department = vcard.Department;
            contact.JobTitle = vcard.Role;
            contact.BirthDate = vcard.BirthDate;
            contact.xv_Salutation = vcard.Gender == vCardGender.Male ? new OptionSetValue(772600000) : (vcard.Gender == vCardGender.Female ? new OptionSetValue(772600001) : null);
            contact.Description += string.Join(" ", vcard.Notes.Select(n => n.Text));

            return contact;
        }

        private Account UpdateAccountWithVcard(IOrganizationService service, Account account, vCard vcard)
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
            account.xv_Land = account.Address1_Country != null ? service.GetCountryReferenceByName(account.Address1_Country) : null;

            account.xv_Firmenklassifizierung = vcard.Categories.Count > 0 ? service.GetCompanyClassificationReferenceByName(vcard.Categories[0]) : null;

            return account;
        }
    }
}