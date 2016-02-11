using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace VcardCrmImporter
{
    using System.Collections.Specialized;
    using System.Web.Mvc;

    using Microsoft.Azure;

    public class MandrillWebhookAttribute : ActionFilterAttribute
    {
        public string Key { get; set; }

        public string KeyAppSetting { get; set; }

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var key = this.Key ?? CloudConfigurationManager.GetSetting(this.KeyAppSetting);

            var signature = this.GenerateSignature(key, filterContext.RequestContext.HttpContext.Request.Url.ToString(), filterContext.RequestContext.HttpContext.Request.Form);
            
            base.OnActionExecuting(filterContext);
        }

        private string GenerateSignature(string key, string url, NameValueCollection form)
        {

            return null;
        }
    }
}