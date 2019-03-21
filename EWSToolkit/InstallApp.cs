using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;

namespace EWS
{
    public class InstallApp
    {
        public static void InstallNewApp(ExchangeService service, string manifest) 
        {
            if (File.Exists(manifest))
            {
                Console.WriteLine(" [>] Using {0}", manifest);

                using (var fs = File.OpenRead(manifest))
                {
                    service.InstallApp(fs);
                }
            }
            else
            {
                Console.WriteLine(" [x] Manifest does not exist");
                return;
            }
                
            Console.WriteLine(" [>] Done");
        }
    }
}
