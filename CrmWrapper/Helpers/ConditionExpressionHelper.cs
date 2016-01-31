namespace CrmWrapper.Helpers
{
    using Microsoft.Xrm.Sdk.Query;

    internal static class ConditionExpressionHelper
    {
        internal static ConditionExpression CreatePublisherIdCondition(int publisherPlatformId)
        {
            return CreateEqualsCondition("xv_platformid", publisherPlatformId);
        }

        internal static ConditionExpression CreateIsPublisherCondition()
        {
            return CreateEqualsCondition("xv_firmenklassifizierung", CrmIds.CompanyTypePublisher);
        }

        internal static ConditionExpression CreatePublisherNameExactCondition(string companyName)
        {
            return CreateEqualsCondition("name", companyName);
        }

        internal static ConditionExpression CreatePublisherNameBeginsWithCondition(string name)
        {
            var condition = new ConditionExpression
                {
                    AttributeName = "name", 
                    Operator = ConditionOperator.BeginsWith
                };

            condition.Values.Add(name);
            return condition;
        }

        internal static ConditionExpression CreatePlatformIdEqualsCondition(int entityPlatformId)
        {
            return CreateEqualsCondition("xv_platformid", entityPlatformId);
        }

        internal static ConditionExpression CreateFirstNameEqualsCondition(string firstName)
        {
            return CreateEqualsCondition("firstname", firstName);
        }

        internal static ConditionExpression CreateLastNameEqualsCondition(string lastName)
        {
            return CreateEqualsCondition("lastname", lastName);
        }

        internal static ConditionExpression CreateCompanyNameCondition(string companyName)
        {
            return CreateEqualsCondition("name", companyName);
        }

        internal static ConditionExpression CreateIndustryCondition(string industryName)
        {
            return CreateEqualsCondition("xv_name", industryName);
        }

        internal static ConditionExpression CreateCountryCondition(string countryName)
        {
            return CreateEqualsCondition("xv_name", countryName);
        }

        internal static ConditionExpression CreateEqualsCondition(string attributeName, object value)
        {
            var condition = new ConditionExpression
            {
                AttributeName = attributeName,
                Operator = ConditionOperator.Equal
            };
            condition.Values.Add(value);
            return condition;
        }
    }
}