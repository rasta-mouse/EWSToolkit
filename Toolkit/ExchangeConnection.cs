using Microsoft.Exchange.WebServices.Data;

using System;
using System.Net;
using System.Security;

namespace Toolkit
{
    public class ExchangeConnection
    {
        public ExchangeService ExchangeService { get; private set; }

        public ExchangeConnection(ExchangeVersion version = ExchangeVersion.Exchange2013_SP1, bool useDefaultCreds = false)
        {
            ExchangeService = new ExchangeService((Microsoft.Exchange.WebServices.Data.ExchangeVersion)version)
            {
                UseDefaultCredentials = useDefaultCreds
            };
        }

        public void SetConnectionCredentials(string username, SecureString password)
        {
            ExchangeService.Credentials = new NetworkCredential(username, password);
        }

        public void SetConnectionCredentials(string username, SecureString password, string domain)
        {
            ExchangeService.Credentials = new NetworkCredential(username, password, domain);
        }

        public void SetConnectionCredentials(string username, string password, string domain = "")
        {
            ExchangeService.Credentials = new NetworkCredential(username, password, domain);
        }

        public void SetAutodiscoverUrl(string emailAddress)
        {
            ExchangeService.AutodiscoverUrl(emailAddress, UrlCallback);
        }

        public void SetAutodiscoverUrl(Uri url)
        {
            ExchangeService.Url = url;
        }

        private bool UrlCallback(string redirectionUrl)
        {
            var uri = new Uri(redirectionUrl);
            return uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase);
        }
    }

    public enum ExchangeVersion
    {
        Exchange2007_SP1 = 0,
        Exchange2010 = 1,
        Exchange2010_SP1 = 2,
        Exchange2010_SP2 = 3,
        Exchange2013 = 4,
        Exchange2013_SP1 = 5
    }
}