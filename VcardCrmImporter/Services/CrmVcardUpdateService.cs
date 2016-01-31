namespace VcardCrmImporter.Services
{
    using System;
    using System.Data.Entity;
    using System.EnterpriseServices;
    using System.Linq;

    using CrmOrganizationClasses;
    using CrmWrapper;

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
            var crmConnection = new CrmConnection(CloudConfigurationManager.GetSetting("CrmConnectionString"));

            using (var serviceProxy = new OrganizationServiceProxy(crmConnection.ServiceUri, crmConnection.HomeRealmUri, crmConnection.ClientCredentials, null))
            {
                // This statement is required to enable early-bound type support.
                serviceProxy.EnableProxyTypes();


                var orgContext = new OrganizationServiceContext(serviceProxy);

                // Retrieve the system user ID of the user to impersonate.
                var impersonatedUserId = await(from user in orgContext.CreateQuery<SystemUser>()
                                          where user.InternalEMailAddress == userEmail 
                                          select user.SystemUserId.Value).FirstOrDefaultAsync();

                // We impersonate the user that has sent the email
                if (impersonatedUserId == new Guid())
                {
                    throw new Exception("User not found in CRM");
                }

                serviceProxy.CallerId = impersonatedUserId;

                var companyNameShortend = CompanyNameHelper.RemoveCommonCompanySuffixes(vcard.Organization);
                var account = await orgContext.CreateQuery<Account>().FirstOrDefaultAsync(a => a.Name.StartsWith(companyNameShortend));
                if (account != null)
                {
                    this.UpdateAccountWithVcard(serviceProxy, account, vcard);
                }
                else
                {
                    account = this.UpdateAccountWithVcard(serviceProxy, null, vcard);
                    account.OwnerId.Id = impersonatedUserId;
                    orgContext.Attach(account);
                }

                orgContext.SaveChanges();

                // Try to get the contact by Firstname, Lastname and Company name
                var contact = await orgContext.CreateQuery<Contact>()
                    .FirstOrDefaultAsync(c => c.FirstName == vcard.GivenName.Trim()
                        && c.LastName == vcard.FamilyName.Trim()
                        && c.ParentCustomerId.Name == vcard.Organization.Trim());

                // Try to search contact by email
                contact = contact ?? await orgContext.CreateQuery<Contact>().FirstOrDefaultAsync(c => c.EMailAddress1 == vcard.EmailAddresses[0].Address.Trim());

                if (contact != null)
                {
                    // If the contact is found, update it
                    this.UpdateContactWithVcard(serviceProxy, contact, vcard);
                    contact.AccountId.Id = account.Id;
                }
                else
                {
                    // If the contact was not found, insert it
                    contact = this.UpdateContactWithVcard(serviceProxy, null, vcard);
                    contact.AccountId.Id = account.Id;
                    contact.OwnerId.Id = impersonatedUserId;
                    orgContext.Attach(contact);
                }

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
            contact.Description += vcard.Notes;

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