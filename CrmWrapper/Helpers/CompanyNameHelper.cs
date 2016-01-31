namespace CrmWrapper.Helpers
{
    using System.Collections.Generic;
    using System.Linq;

    public class CompanyNameHelper
    {
        private static readonly List<string> CompanySuffixes = new[]
            {
                "gmbh & co kg", 
                "gmbh", 
                "mbh", 
                "kgaa", 
                "se", 
                "ag", 
                "gbr", 
                "kg", 
                "& co", 
                "inc", 
                "ltd", 
                "ug", 
                "ug (haftungsbeschränkt)", 
                "ug haftungsbeschränkt", 
                "ohg", 
                "ggmbh", 
                "sa", 
                "sarl", 
                "sl", 
                "srl", 
                "ab", 
                "corp"
            }.OrderByDescending(x => x.Length).ToList();

        public static string RemoveCommonCompanySuffixes(string companyName)
        {
            var name = companyName.ToLowerInvariant().Trim('.');

            foreach (var suffix in CompanySuffixes.Select(s => " " + s))
            {
                if (name.EndsWith(suffix))
                {
                    name = name.Substring(0, name.Length - suffix.Length);
                }
            }

            return name;
        }
    }
}