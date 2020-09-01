# EWSToolkit

The EWS Toolkit is a .NET Framework Library that wraps around the EWS Managed API.  It contains easy methods and overloads to abuse various EWS functionality in C# or PowerShell tooling.

## Example

```
using Toolkit;

namespace DemoApp
{
    class Program
    {
        static void Main(string[] args)
        {
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