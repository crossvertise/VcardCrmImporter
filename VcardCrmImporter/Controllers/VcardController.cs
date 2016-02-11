namespace VcardCrmImporter.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Mvc;

    using Mandrill.Models;
    using Mandrill.Requests.Messages;

    using Microsoft.Azure;

    using Newtonsoft.Json;

    using Thought.vCards;

    using VcardCrmImporter.Services;

    public class VcardController : Controller
    {
        private string mandrillApiKey = CloudConfigurationManager.GetSetting("CrmConnectionString");

        private string errorEmailRecipient = CloudConfigurationManager.GetSetting("ErrorEmailRecipient");

        private string emailSendingAddress = CloudConfigurationManager.GetSetting("EmailSendingAddress");

        // GET: Vcard
        [HttpGet]
        public ActionResult PostVcards()
        {
            return this.View();
        }

        [HttpPost]
        public async Task<ActionResult> PostVcards(HttpPostedFileBase file)
        {
            if (file == null)
            {
                ModelState.AddModelError("file", "An image file must be chosen.");
                return this.View();
            }
            
            var vcardReader = new Thought.vCards.vCardStandardReader();
            using (var streamreader = new StreamReader(file.InputStream))
            {
                var vcard = vcardReader.Read(streamreader);
                var service = new CrmVcardUpdateService();
                await service.UpdateContactWithVcard(vcard, "m.balbach@crossvertise.com");
            }
            

            return this.View("Success");
        }

        [HttpPost]
        [ValidateInput(false)]
        public async Task<ActionResult> MailWebhook()
        {
            var mandrill = new Mandrill.MandrillApi(this.mandrillApiKey);
            string sender = null;

            try
            {
                string validJson = HttpContext.Request.Form["mandrill_events"].Replace("mandrill_events=", string.Empty);

                var webhookEvents = JsonConvert.DeserializeObject<List<WebHookEvent>>(validJson);

                var results = new List<string>();

                foreach (var webhookEvent in webhookEvents)
                {
                    sender = webhookEvent.Msg.Sender;
                    var attachments = webhookEvent.Msg.Attachments;

                    foreach (var attachment in attachments)
                    {
                        if ((attachment.Value.Type != "text/vcard" && attachment.Value.Type != "text/x-vcard") || !attachment.Key.EndsWith(".vcf"))
                        {
                            results.Add(attachment.Key + " (mime-type: " + attachment.Value.Type + "): not imported");
                            continue;
                        }

                        var bytes = attachment.Value.Base64 ? Convert.FromBase64String(attachment.Value.Content) : Encoding.UTF8.GetBytes(attachment.Value.Content);
                        var memoryStream = new MemoryStream(bytes);

                        var vcardReader = new vCardStandardReader();
                        using (var streamreader = new StreamReader(memoryStream))
                        {
                            var vcard = vcardReader.Read(streamreader);
                            var service = new CrmVcardUpdateService();
                            await service.UpdateContactWithVcard(vcard, sender);

                            results.Add(string.Format("Processed {0} successfully", attachment.Key));
                        }
                    }
                }

                var email = new EmailMessage
                {
                    FromEmail = this.emailSendingAddress,
                    To = new List<EmailAddress> { new EmailAddress(sender ?? this.errorEmailRecipient) },
                    Subject = "VCard import completed",
                    Text = string.Format("{0} files processed, {1} contacts imported \r\n", results.Count, results.Count(r => r.Contains("successfully"))) + string.Join("\n", results),
                };

                await mandrill.SendMessage(new SendMessageRequest(email));

                return this.View((object)validJson);
            }
            catch (Exception ex)
            {
                var email = new EmailMessage
                                {
                                    FromEmail = this.emailSendingAddress,
                                    To = new List<EmailAddress> { new EmailAddress(sender ?? this.errorEmailRecipient) },
                                    Subject = "Error in VCard Import",
                                    Text = JsonConvert.SerializeObject(ex)
                                };

                mandrill.SendMessage(new SendMessageRequest(email));

                throw;
            }
        }
    }
}