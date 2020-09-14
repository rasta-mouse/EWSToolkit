# EWSToolkit

The EWS Toolkit is a .NET Framework Library that wraps around the EWS Managed API.  It contains easy methods and overloads to abuse various EWS functionality in C# or PowerShell tooling.

## Example

Auth and send email
```
using Toolkit;

namespace DemoApp
{
    class Program
    {
        static void Main(string[] args)
        {
             //Program.exe random@example.org Passw0rd123
             
            var connection = new ExchangeConnection(ExchangeVersion.Exchange2013_SP1);
            connection.SetConnectionCredentials(args[0], args[1].ToSecureString());
            connection.SetAutodiscoverUrl(args[0]);

            var mail = new Email(connection);
            mail.AddRecipient("rasta@rastamouse.me");
            mail.SetSubject("Test Email");
            mail.SetBodyFromText("This is a test email");
            mail.Send();
        }
    }
}
```

Auth and search for keywords in the inbox and sent folders
```
using Toolkit;

namespace DemoApp
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //Program.exe random@example.org Passw0rd123

            ExchangeConnection connection = new ExchangeConnection(ExchangeVersion.Exchange2013_SP1);
            connection.SetConnectionCredentials(args[0], args[1].ToSecureString());
            connection.SetAutodiscoverUrl(args[0]);

            Search searchData = new Search(connection);

            //Add single search terms
            searchData.AddBodySearchTerm("Administrator");
            searchData.AddSubjectSearchTerm("Support");

            //Add multi search terms
            searchData.AddBodySearchTerms(new string[] { "Password", "Login" });
            searchData.AddSubjectSearchTerms(new string[] { "Onboarding", "Welcome" });

            List<SearchResult> searchResult = searchData.SearchFolder(FolderName.Inbox | FolderName.SentItems);
        }
    }
}
```
