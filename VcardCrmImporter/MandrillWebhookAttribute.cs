using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace VcardCrmImporter
{
    using System.Collections.Specialized;
    using System.Security.Cryptography;
    using System.Web.Mvc;

    using Microsoft.Azure;

    public class MandrillWebhookAttribute : ActionFilterAttribute
    {
        public string Key { get; set; }

        public string KeyAppSetting { get; set; }

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var key = this.Key ?? CloudConfigurationManager.GetSetting(this.KeyAppSetting);

            var signature = this.GenerateSignature(key, filterContext.HttpContext.Request.Url.ToString(), filterContext.RequestContext.HttpContext.Request.Form);

            var mandrillSignature = filterContext.HttpContext.Request.Headers.GetValues("X-Mandrill-Signature").FirstOrDefault();

            if (mandrillSignature == null || mandrillSignature != signature)
            {
                throw new HttpException(401, "The webhook call was not properly authorized.");
            }

            base.OnActionExecuting(filterContext);
        }

        private string GenerateSignature(string key, string url, NameValueCollection form)
        {
            var sourceString = url;
            foreach (var formKey in form.AllKeys)
            {
                sourceString += formKey + form[formKey];
            }

            byte[] byteKey = System.Text.Encoding.UTF8.GetBytes(key);
            byte[] byteValue = System.Text.Encoding.UTF8.GetBytes(sourceString);
            HMACSHA1 myhmacsha1 = new HMACSHA1(byteKey);
            byte[] hashValue = myhmacsha1.ComputeHash(byteValue);
            string generatedSignature = Convert.ToBase64String(hashValue);

            return generatedSignature;
        }
    }
}