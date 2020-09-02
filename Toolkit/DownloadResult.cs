using System.Collections.Generic;

namespace Toolkit
{
    public class DownloadResult
    {
        public string Body { get; set; }
        public List<Attachment> Attachments { get; set; } = new List<Attachment>();
    }

    public class Attachment
    {
        public string Filename { get; set; }
        public byte[] Content { get; set; }
    }
}