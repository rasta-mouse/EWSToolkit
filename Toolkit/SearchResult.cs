using System;

namespace Toolkit
{
    public class SearchResult
    {
        public string UniqueId { get; set; }
        public DateTime ReceivedOn { get; set; }
        public string Subject { get; set; }
        public bool HasAttachments { get; set; }
    }
}