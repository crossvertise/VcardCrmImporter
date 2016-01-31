namespace VcardCrmImporter.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Mvc;

    using CrmOrganizationClasses;

    using VcardCrmImporter.Services;

    public class VcardController : Controller
    {
        // GET: Vcard
        [HttpGet]
        public ActionResult PostVcards()
        {
            return this.View();
        }

        [HttpPost]
        public async Task<ActionResult> PostVcards(IEnumerable<HttpPostedFileBase> files)
        {
            foreach (var file in files)
            {
                var vcardReader = new Thought.vCards.vCardStandardReader();
                using (var streamreader = new StreamReader(file.InputStream))
                {
                    var vcard = vcardReader.Read(streamreader);
                    var service = new CrmVcardUpdateService();
                    await service.UpdateContactWithVcard(vcard, "m.balbach@crossvertise.com");
                }
            }

            return new HttpStatusCodeResult(200);
        }

        public ActionResult MailWebhook()
        {   
            return new HttpStatusCodeResult(200);
        }
    }
}