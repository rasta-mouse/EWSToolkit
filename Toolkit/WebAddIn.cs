using Microsoft.Exchange.WebServices.Data;

using System;
using System.IO;

namespace Toolkit
{
    public class WebAddIn
    {
        ExchangeService Service;
        string ManifestPath;

        public WebAddIn(ExchangeConnection connection)
        {
            Service = connection.ExchangeService;
        }

        public void AddManifest(string path)
        {
            if (!File.Exists(path))
            {
                throw new ArgumentException("File not found");
            }

            ManifestPath = path;
        }

        public void InstallApp()
        {
            using (var fs = File.OpenRead(ManifestPath))
            {
                Service.InstallApp(fs);
            }
        }
    }
}
