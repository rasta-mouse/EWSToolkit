using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace EWS
{
    public class Email
    {
        public static void Send(ExchangeService service, List<string> recipients, string subject, string format, string template, string attachment, string from)
        {
            Console.WriteLine(" [>] Constructing email");

            EmailMessage message = new EmailMessage(service);

            foreach (string recipient in recipients)
            {
                Console.WriteLine(" [>] Adding recipient: {0}", recipient);
                message.ToRecipients.Add(recipient);
            }

            Console.WriteLine(" [>] Subject: {0}", subject);
            message.Subject = subject;

            Console.WriteLine(" [>] Template: {0}", template);
            message.Body = File.ReadAllText(template);

            if (format == "txt")
            {
                Console.WriteLine(" [>] Format: Text");
                message.Body.BodyType = BodyType.Text;
            }
            else
            {
                Console.WriteLine(" [>] Format: HTML");
                message.Body.BodyType = BodyType.HTML;
            }

            if (!string.IsNullOrEmpty(attachment))
            {
                Console.WriteLine(" [>] Attaching: {0}", attachment);
                message.Attachments.AddFileAttachment(attachment);
            }

            if (!string.IsNullOrEmpty(from))
            {
                Console.WriteLine(" [>] Setting from address: {0}", from);
                message.From = from;
            }

            Console.WriteLine(" [>] Sending...");
            message.Send();
            Console.WriteLine(" [>] Done");

        }
    }
}
