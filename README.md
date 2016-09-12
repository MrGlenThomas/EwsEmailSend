# EwsEmailSend
A simple .NET console application to send email messages using Exchange Web Services

'Send' - Send an email message

Expected usage: ExchangeEmailSend.exe Send <options>
<options> available:
  -d, --domain=VALUE         The exchange server domain.
  -u, --user=VALUE           The network credential username.
  -p, --password=VALUE       The network credential password.
  -e, --email=VALUE          The email address to send from.
  -r, --recipient=VALUE      The email address of the recipient.
  -s, --subject=VALUE        The email subject.
  -b, --body=VALUE           The file path of the HTML body.
      --url                  The exchange server URL
