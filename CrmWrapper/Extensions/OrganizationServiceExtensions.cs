namespace CrmWrapper.Extensions
{
    using System.Collections.Generic;
    using System.Linq;

    using CrmOrganizationClasses;

    using CrmWrapper.Exceptions;
    using CrmWrapper.Helpers;

    using Microsoft.Practices.Unity.Utility;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    public static class OrganizationServiceExtensions
    {
        public static EntityReference GetStandardPriceLevelReference(this IOrganizationService service)
        {
            var query = new QueryExpression(PriceLevel.EntityLogicalName);

            var standardPriceLevel = service.RetrieveMultiple(query)
                .Entities
                .First();

            return new EntityReference(PriceLevel.EntityLogicalName, standardPriceLevel.Id);
        }

        public static Contact GetContactByPlatformId(this IOrganizationService service, int personId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePlatformIdEqualsCondition(personId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(Contact.EntityLogicalName)
                {
                    ColumnSet = new ColumnSet("parentcustomerid")
                };

            query.Criteria.AddFilter(filter);

            var contact = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (contact == null)
            {
                throw new CrmEntityNotFoundException("Contact not found.");
            }

            return (Contact)contact;
        }

        public static SalesOrder GetSalesOrderByPlatformId(this IOrganizationService service, int salesOrderId)
        {
            // The double "t" is important in "xv_plattformid" here, that's why we do not use ConditionExpressionHelper.CreatePlatformIdEqualsCondition
            var platformIdCondition = new ConditionExpression
                {
                    AttributeName = "xv_plattformid",
                    Operator = ConditionOperator.Equal
                };

            platformIdCondition.Values.Add(salesOrderId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(SalesOrder.EntityLogicalName);

            query.Criteria.AddFilter(filter);

            var salesOrder = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (salesOrder == null)
            {
                throw new CrmEntityNotFoundException("SalesOrder not found.");
            }

            return (SalesOrder)salesOrder;
        }

        public static EntityReference GetSalesOrderItemReferenceByPlatformId(this IOrganizationService service, int salesOrderItemId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePlatformIdEqualsCondition(salesOrderItemId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(SalesOrderDetail.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var salesOrderItem = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (salesOrderItem == null)
            {
                throw new CrmEntityNotFoundException("SalesOrderItem not found.");
            }

            return new EntityReference(SalesOrderDetail.EntityLogicalName, salesOrderItem.Id);
        }

        public static IEnumerable<EntityReference> GetSalesOrderItemsByOrder(this IOrganizationService service, EntityReference orderReference)
        {
            var salesOrderIdCondition = ConditionExpressionHelper.CreateEqualsCondition("salesorderid", orderReference.Id);

            var filter = new FilterExpression();
            filter.Conditions.Add(salesOrderIdCondition);

            var query = new QueryExpression(SalesOrderDetail.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            return service.RetrieveMultiple(query)
                .Entities
                .Select(item => item.ToEntityReference());
        }

        public static Opportunity GetOpportunityByPlatformId(this IOrganizationService service, int opportunityId)
        {
            // The double "t" is important in "xv_plattformid" here, that's why we do not use ConditionExpressionHelper.CreatePlatformIdEqualsCondition
            var platformIdCondition = new ConditionExpression
            {
                AttributeName = "xv_plattformid",
                Operator = ConditionOperator.Equal
            };

            platformIdCondition.Values.Add(opportunityId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(Opportunity.EntityLogicalName);

            query.Criteria.AddFilter(filter);

            var salesOrder = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (salesOrder == null)
            {
                throw new CrmEntityNotFoundException("Opportunity not found.");
            }

            return (Opportunity)salesOrder;
        }

        public static EntityReference GetOpportunityItemReferenceByPlatformId(this IOrganizationService service, int opportunityItemId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePlatformIdEqualsCondition(opportunityItemId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(OpportunityProduct.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var opportunityItem = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (opportunityItem == null)
            {
                throw new CrmEntityNotFoundException("OpportunityItem not found.");
            }

            return new EntityReference(OpportunityProduct.EntityLogicalName, opportunityItem.Id);
        }

        public static IEnumerable<EntityReference> GetOpportunityItemsByOpportunityId(this IOrganizationService service, EntityReference opportunityReference)
        {
            var opportunityIdCondition = ConditionExpressionHelper.CreateEqualsCondition("opportunityid", opportunityReference.Id);

            var filter = new FilterExpression();
            filter.Conditions.Add(opportunityIdCondition);

            var query = new QueryExpression(OpportunityProduct.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            return service.RetrieveMultiple(query)
                .Entities
                .Select(item => item.ToEntityReference());
        }

        public static EntityReference GetPublisherReferenceByPlatformId(this IOrganizationService service, int publisherId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePublisherIdCondition(publisherId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(Account.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var account = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (account == null)
            {
                throw new CrmEntityNotFoundException("Account not found with platform id: " + publisherId);
            }

            return new EntityReference(Account.EntityLogicalName, account.Id);
        }

        public static Quote GetQuoteByPlatformId(this IOrganizationService service, int quoteId, IEnumerable<string> columnsToLoad = null)
        {
            // The double "t" is important in "xv_plattformid" here, that's why we do not use ConditionExpressionHelper.CreatePlatformIdEqualsCondition
            var platformIdCondition = new ConditionExpression
            {
                AttributeName = "xv_plattformid",
                Operator = ConditionOperator.Equal
            };

            platformIdCondition.Values.Add(quoteId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = CreateQuery(Quote.EntityLogicalName, filter, columnsToLoad);

            var quote = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (quote == null)
            {
                throw new CrmEntityNotFoundException("Quote not found.");
            }

            return (Quote)quote;
        }

        public static EntityReference GetQuoteItemReferenceByPlatformId(this IOrganizationService service, int quoteItemId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePlatformIdEqualsCondition(quoteItemId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(QuoteDetail.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var quoteDetail = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (quoteDetail == null)
            {
                throw new CrmEntityNotFoundException("QuoteDetail not found.");
            }

            return new EntityReference(QuoteDetail.EntityLogicalName, quoteDetail.Id);
        }

        public static IEnumerable<EntityReference> GetQuoteItemsByQuoteId(this IOrganizationService service, EntityReference quoteReference)
        {
            var quoteIdCondition = ConditionExpressionHelper.CreateEqualsCondition("quoteid", quoteReference.Id);

            var filter = new FilterExpression();
            filter.Conditions.Add(quoteIdCondition);

            var query = new QueryExpression(QuoteDetail.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            return service.RetrieveMultiple(query)
                .Entities
                .Select(item => item.ToEntityReference());
        }

        public static EntityReference GetPublisherReferenceByName(this IOrganizationService service, string companyName)
        {
            try
            {
                return service.GetPublisherReferenceByNameExact(companyName);
            }
            catch (CrmEntityNotFoundException)
            {
                return service.GetPublisherReferenceByNameFuzzy(companyName);
            }
        }

        public static EntityReference GetPublisherReferenceByNameExact(this IOrganizationService service, string companyName)
        {
            var nameCondition = ConditionExpressionHelper.CreatePublisherNameExactCondition(companyName);
            var isPublisherCondition = ConditionExpressionHelper.CreateIsPublisherCondition();

            var filter = new FilterExpression();
            filter.Conditions.Add(nameCondition);
            filter.Conditions.Add(isPublisherCondition);

            var query = new QueryExpression(Account.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var account = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (account == null)
            {
                throw new CrmEntityNotFoundException("Account not found with name: " + companyName);
            }

            return new EntityReference(Account.EntityLogicalName, account.Id);
        }

        public static EntityReference GetPublisherReferenceByNameFuzzy(this IOrganizationService service, string companyName)
        {
            var name = CompanyNameHelper.RemoveCommonCompanySuffixes(companyName);
            var nameCondition = ConditionExpressionHelper.CreatePublisherNameBeginsWithCondition(name);
            var isPublisherCondition = ConditionExpressionHelper.CreateIsPublisherCondition();

            var filter = new FilterExpression();
            filter.Conditions.Add(nameCondition);
            filter.Conditions.Add(isPublisherCondition);

            var query = new QueryExpression(Account.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var accounts = service.RetrieveMultiple(query).Entities;

            if (accounts.Count() > 1)
            {
                throw new CrmAmbiguousMatchException(
                    string.Format("Found multiple fuzzy matches when searching for {0}.  Fuzzy match search: {1}", companyName, name));
            }

            var account = accounts.FirstOrDefault();
            if (account == null)
            {
                throw new CrmEntityNotFoundException("Account not found with name: " + companyName);
            }

            return new EntityReference(Account.EntityLogicalName, account.Id);
        }

        public static EntityReference GetAccountReferenceByNameFuzzy(this IOrganizationService service, string companyName)
        {
            var name = CompanyNameHelper.RemoveCommonCompanySuffixes(companyName);
            var nameCondition = ConditionExpressionHelper.CreatePublisherNameBeginsWithCondition(name);

            var filter = new FilterExpression();
            filter.Conditions.Add(nameCondition);

            var query = new QueryExpression(Account.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var accounts = service.RetrieveMultiple(query).Entities;

            if (accounts.Count() > 1)
            {
                throw new CrmAmbiguousMatchException(
                    string.Format("Found multiple fuzzy matches when searching for {0}.  Fuzzy match search: {1}", companyName, name));
            }

            var account = accounts.FirstOrDefault();
            if (account == null)
            {
                throw new CrmEntityNotFoundException("Account not found with name: " + companyName);
            }

            return new EntityReference(Account.EntityLogicalName, account.Id);
        }

        public static EntityReference GetManagerReferenceByName(this IOrganizationService service, string firstName, string lastName)
        {
            var firstNameCondition = ConditionExpressionHelper.CreateFirstNameEqualsCondition(firstName);
            var lastNameCondition = ConditionExpressionHelper.CreateLastNameEqualsCondition(lastName);

            var filter = new FilterExpression();
            filter.Conditions.Add(firstNameCondition);
            filter.Conditions.Add(lastNameCondition);

            var query = new QueryExpression(SystemUser.EntityLogicalName)
                            {
                                ColumnSet = new ColumnSet(true)
                            };

            query.Criteria.AddFilter(filter);

            var users = service.RetrieveMultiple(query)
                .Entities.ToList();

            if (!users.Any())
            {
                throw new CrmEntityNotFoundException(string.Format("User '{0} {1}' was not found.", firstName, lastName));
            }

            if (users.Count > 1)
            {
                throw new CrmAmbiguousMatchException(string.Format("{0} users found with name '{1} {2}'.", users.Count, firstName, lastName));
            }

            return users.Single().ToEntityReference();
        }

        public static EntityReference GetCompanyReferenceByName(this IOrganizationService service, string companyName)
        {
            var compayNameCondition = ConditionExpressionHelper.CreateCompanyNameCondition(companyName);
            var filter = new FilterExpression();
            filter.Conditions.Add(compayNameCondition);

            var query = new QueryExpression(Account.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true)
            };

            query.Criteria.AddFilter(filter);

            var companies = service.RetrieveMultiple(query)
                .Entities.ToList();

            if (!companies.Any())
            {
                throw new CrmEntityNotFoundException(string.Format("Company '{0}' was not found.", companyName));
            }

            if (companies.Count > 1)
            {
                throw new CrmAmbiguousMatchException(string.Format("{0} companies found with name '{1}'.", companies.Count, companyName));
            }
            
            return companies.Single().ToEntityReference();
        }

        public static EntityReference GetCompanyReferenceById(this IOrganizationService service, string companyName)
        {
            var compayNameCondition = ConditionExpressionHelper.CreateCompanyNameCondition(companyName);
            var filter = new FilterExpression();
            filter.Conditions.Add(compayNameCondition);

            var query = new QueryExpression(Account.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true)
            };

            query.Criteria.AddFilter(filter);

            var companies = service.RetrieveMultiple(query)
                .Entities.ToList();

            if (!companies.Any())
            {
                throw new CrmEntityNotFoundException(string.Format("Company '{0}' was not found.", companyName));
            }

            if (companies.Count > 1)
            {
                throw new CrmAmbiguousMatchException(string.Format("{0} companies found with name '{1}'.", companies.Count, companyName));
            }

            return companies.Single().ToEntityReference();
        }

        public static EntityReference GetIndustryReferenceByName(this IOrganizationService service, string industryName)
        {
            if (string.IsNullOrEmpty(industryName))
            {
                return default(EntityReference);
            }

            var industryCondition = ConditionExpressionHelper.CreateIndustryCondition(industryName);
            var filter = new FilterExpression();
            filter.Conditions.Add(industryCondition);

            var query = new QueryExpression(xv_branche.EntityLogicalName);

            query.Criteria.AddFilter(filter);

            var companies = service.RetrieveMultiple(query)
                .Entities.ToList();

            if (!companies.Any())
            {
                return default(EntityReference);
            }

            return companies.First().ToEntityReference();
        }

        public static EntityReference GetCountryReferenceByName(this IOrganizationService service, string countryName)
        {
            if (string.IsNullOrWhiteSpace(countryName))
            {
                return default(EntityReference);
            }

            var countryCondition = ConditionExpressionHelper.CreateCountryCondition(countryName);
            var filter = new FilterExpression();
            filter.Conditions.Add(countryCondition);

            var query = new QueryExpression(xv_land.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var countries = service.RetrieveMultiple(query)
                .Entities.ToList();

            if (!countries.Any())
            {
                return default(EntityReference);
            }

            return countries.First().ToEntityReference();
        }

        public static EntityReference GetProductTypeReferenceByName(this IOrganizationService service, string productTypeName)
        {
            Guard.ArgumentNotNullOrEmpty(productTypeName, "productTypeName");

            var productTypeNameCondition = ConditionExpressionHelper.CreateEqualsCondition("xv_name", productTypeName);
            var filter = new FilterExpression();
            filter.Conditions.Add(productTypeNameCondition);

            var query = new QueryExpression(xv_gattungen.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var productTypeEntity = service.RetrieveMultiple(query)
                .Entities
                .FirstOrDefault();

            if (productTypeEntity == null)
            {
                throw new CrmEntityNotFoundException(string.Format("Product type '{0}' was not found.", productTypeName));
            }

            return productTypeEntity.ToEntityReference();
        }

        public static EntityReference GetCompanyClassificationReferenceByName(this IOrganizationService service, string companyClassification)
        {
            Guard.ArgumentNotNullOrEmpty(companyClassification, "companyClassification");

            var productTypeNameCondition = ConditionExpressionHelper.CreateEqualsCondition("xv_name", companyClassification);
            var filter = new FilterExpression();
            filter.Conditions.Add(productTypeNameCondition);

            var query = new QueryExpression(xv_firmenklassifizierung.EntityLogicalName);
            query.Criteria.AddFilter(filter);

            var productTypeEntity = service.RetrieveMultiple(query)
                .Entities
                .FirstOrDefault();

            if (productTypeEntity == null)
            {
                throw new CrmEntityNotFoundException(string.Format("Company classification '{0}' was not found.", companyClassification));
            }

            return productTypeEntity.ToEntityReference();
        }

        public static xv_platformmessage GetMessageByPlatformId(this IOrganizationService service, int messageId)
        {
            var platformIdCondition = ConditionExpressionHelper.CreatePlatformIdEqualsCondition(messageId);

            var filter = new FilterExpression();
            filter.Conditions.Add(platformIdCondition);

            var query = new QueryExpression(xv_platformmessage.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true)
            };

            query.Criteria.AddFilter(filter);

            var message = service.RetrieveMultiple(query).Entities
                .FirstOrDefault();

            if (message == null)
            {
                throw new CrmEntityNotFoundException("Message not found.");
            }

            return (xv_platformmessage)message;
        }

        private static QueryExpression CreateQuery(string entityLogicalName, FilterExpression filter, IEnumerable<string> propertiesToRead = null)
        {
            var query = new QueryExpression(entityLogicalName)
                {
                    ColumnSet = new ColumnSet()
                };

            if (propertiesToRead != null)
            {
                query.ColumnSet.AddColumns(propertiesToRead.Select(p => p).ToArray());
            }

            if (filter != null)
            {
                query.Criteria.AddFilter(filter);
            }

            return query;
        }
    }
}
