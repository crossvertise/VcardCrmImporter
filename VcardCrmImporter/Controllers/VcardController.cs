namespace VcardCrmImporter.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
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
        private string mandrillApiKey = CloudConfigurationManager.GetSetting("MandrillApiKey");

        private string errorEmailRecipient = CloudConfigurationManager.GetSetting("ErrorEmailRecipient");

        private string emailSendingAddress = CloudConfigurationManager.GetSetting("EmailSendingAddress");

        public async Task<ActionResult> MailTest()
        {
            var mandrill = new Mandrill.MandrillApi(this.mandrillApiKey);
            var email = new EmailMessage
            {
                FromEmail = this.emailSendingAddress,
                To = new List<EmailAddress> { new EmailAddress(this.errorEmailRecipient) },
                Subject = "VCard import finished",
                Text = "sdfsdf",
            };

            await mandrill.SendMessage(new SendMessageRequest(email));

            return new HttpStatusCodeResult(200);
        }


        [HttpPost]
        [ValidateInput(false)]
        [MandrillWebhook(KeyAppSetting = "MandrillWebhookKey")]
        public async Task<ActionResult> MailWebhook()
        {
            // Initialize a mandrill client for sending emails
            var mandrill = new Mandrill.MandrillApi(this.mandrillApiKey);
            string sender = null;

            try
            {
                // Get the JSON payload
                string validJson = HttpContext.Request.Form["mandrill_events"].Replace("mandrill_events=", string.Empty);

                if (string.IsNullOrWhiteSpace(validJson))
                {
                    return new HttpStatusCodeResult(400);
                }

                var webhookEvents = JsonConvert.DeserializeObject<List<WebHookEvent>>(validJson);
                if (webhookEvents == null)
                {
                    return new HttpStatusCodeResult(400);
                }

                var results = new List<string>();

                foreach (var webhookEvent in webhookEvents)
                {
                    sender = webhookEvent.Msg.FromEmail;
                    var attachments = webhookEvent.Msg.Attachments;

                    // Verify the security by checking authorized senders
                    if (webhookEvent.Msg.Spf.Result != "pass" && webhookEvent.Msg.Dkim.Valid != true)
                    {
                        throw new Exception("Sender not authorized: SPF / DKIM Check failed.");
                    }

                    // Create an instance of the CRM service that acts in the name of the sender of the email
                    using (var service = new CrmVcardUpdateService(sender))
                    {
                        // Loop through the attachments
                        foreach (var attachment in attachments)
                        {
                            // Get the VCards
                            var filename = this.DecodeBase64Names(attachment.Key);

                            if ((attachment.Value.Type != "text/vcard" && attachment.Value.Type != "text/x-vcard") || !filename.EndsWith(".vcf"))
                            {
                                results.Add(string.Format("{0}: not imported (mime-type: {1})", filename, attachment.Value.Type));
                                continue;
                            }

                            var bytes = attachment.Value.Base64 ? Convert.FromBase64String(attachment.Value.Content) : Encoding.UTF8.GetBytes(attachment.Value.Content);
                            var memoryStream = new MemoryStream(bytes);
                            var vcardReader = new vCardStandardReader();
                            using (var streamreader = new StreamReader(memoryStream))
                            {
                                var vcard = vcardReader.Read(streamreader);

                                // Import the VCards into the CRM
                                var result = service.UpdateContactWithVcard(vcard, filename);
                                results.Add(result);
                            }
                        }
                    }
                }

                // Sends a mail with the results
                var email = new EmailMessage
                {
                    FromEmail = this.emailSendingAddress,
                    To = new List<EmailAddress> { new EmailAddress(sender ?? this.errorEmailRecipient) },
                    Subject = "VCard import finished",
                    Text = "Results:\t\n\t\n" + string.Join("\t\n", results),
                    Html = "Results:<br/>\t\n <br/>\t\n" + string.Join("<br/>\t\n", results),
                };

                await mandrill.SendMessage(new SendMessageRequest(email));

                return this.View((object)validJson);
            }
            catch (Exception ex)
            {
                // Send a mail with the error
                var email = new EmailMessage
                                {
                                    FromEmail = this.emailSendingAddress,
                                    To = new List<EmailAddress> { new EmailAddress(sender ?? this.errorEmailRecipient) },
                                    Subject = "Error in VCard Import",
                                    Text = JsonConvert.SerializeObject(ex, Formatting.Indented)
                                };

                await mandrill.SendMessage(new SendMessageRequest(email));

                throw;
            }
        }

        private string DecodeBase64Names(string input)
        {
            var match = Regex.Match(input, @"^=\?utf-8\?B\?(?<base64>.+)\?=$");
            if (match.Success)
            {
                byte[] data = Convert.FromBase64String(match.Groups["base64"].Value);
                return Encoding.UTF8.GetString(data);
            }

            return input;
        }
    }
}