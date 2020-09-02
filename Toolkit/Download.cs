using Microsoft.Exchange.WebServices.Data;

using System.Collections.Generic;
using System.IO;

namespace Toolkit
{
    public class Download
    {
        ExchangeService Service;

        public Download(ExchangeConnection connection)
        {
            Service = connection.ExchangeService;
        }

        public List<DownloadResult> DownloadEmailMessages(string[] ids, bool includeAttachments = true)
        {
            var result = new List<DownloadResult>();

            foreach (var id in ids)
            {
                result.Add(DownloadEmailMessage(id, includeAttachments));
            }

            return result;
        }

        public DownloadResult DownloadEmailMessage(string id, bool includeAttachments = true)
        {
            var result = new DownloadResult();

            if (includeAttachments)
            {
                var email = Item.Bind(Service, id, new PropertySet(ItemSchema.Body, ItemSchema.Attachments));
                result.Body = email.Body.Text;

                if (email.HasAttachments)
                {
                    foreach (FileAttachment file in email.Attachments)
                    {
                        result.Attachments.Add(new Attachment
                        {
                            Filename = file.Name,
                            Content = DownloadAttachment(file)
                        });
                    }
                }
            }
            else
            {
                var email = Item.Bind(Service, id, new PropertySet(ItemSchema.Body));
                result.Body = email.Body.Text;
            }

            return result;
        }

        private byte[] DownloadAttachment(FileAttachment attachment)
        {
            using (var fs = new MemoryStream())
            {
                attachment.Load(fs);
                return fs.ToArray();
            }
        }
    }
}