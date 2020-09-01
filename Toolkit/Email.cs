using Microsoft.Exchange.WebServices.Data;

using System;
using System.IO;

namespace Toolkit
{
    public class Email
    {
        EmailMessage Message;

        public Email(ExchangeConnection connection, BodyType type = BodyType.Text)
        {
            Message = new EmailMessage(connection.ExchangeService)
            {
                Body = ""
            };

            Message.Body.BodyType = (Microsoft.Exchange.WebServices.Data.BodyType)type;
        }

        public void AddRecipients(string[] recipients)
        {
            foreach (var recipient in recipients)
            {
                AddRecipient(recipient);
            }
        }

        public void AddRecipient(string recipient)
        {
            Message.ToRecipients.Add(recipient);
        }

        public void SetSubject(string subject)
        {
            Message.Subject = subject;
        }

        public void SetBodyFromTemplate(string path)
        {
            if (!File.Exists(path))
            {
                throw new ArgumentException("File not found");
            }

            Message.Body = File.ReadAllText(path);
        }

        public void SetBodyFromText(string text)
        {
            Message.Body = text;
        }

        public void SetFromAddress(string address)
        {
            Message.From = address;
        }

        public void AddAttachment(string path)
        {
            if (!File.Exists(path))
            {
                throw new ArgumentException("File not found");
            }

            Message.Attachments.AddFileAttachment(path);
        }

        public void Send()
        {
            Message.Send();
        }
    }

    public enum BodyType
    {
        HTML = 0,
        Text = 1
    }
}