
namespace ExchangeEmailSend
{
    using System;
    using System.IO;

    using ManyConsole;

    using Microsoft.Exchange.WebServices.Data;

    class SendCommand : ConsoleCommand
    {
        public string Recipient { get; set; }

        public string Subject { get; set; }

        public string BodyFilePath { get; set; }

        public string Domain { get; set; }

        public string UserName { get; set; }

        public string Password { get; set; }

        public string ExchangeServerUrl { get; set; }

        public string SenderEmail { get; set; }

        public SendCommand()
        {
            IsCommand("Send", "Send an email message");
            HasRequiredOption("d|domain=", "The exchange server domain.", d => Domain = d);
            HasRequiredOption("u|user=", "The network credential username.", u => UserName = u);
            HasRequiredOption("p|password=", "The network credential password.", p => Password = p);
            HasRequiredOption("e|email=", "The email address to send from.", e => SenderEmail = e);
            HasRequiredOption("r|recipient=", "The email address of the recipient.", r => Recipient = r);
            HasRequiredOption("s|subject=", "The email subject.", s => Subject = s);
            HasRequiredOption("b|body=", "The file path of the HTML body.", b => BodyFilePath = b);
            HasOption("url", "The exchange server URL", url => ExchangeServerUrl = url);
        }
        
        public override int Run(string[] remainingArguments)
        {
            Console.WriteLine("Reading body file...");
            if (!File.Exists(BodyFilePath))
            {
                Console.WriteLine("Error: body file path not found");
                return 1;
            }
            var body = File.ReadAllText(BodyFilePath);

            Console.WriteLine("Creating service...");
            var service = ExchangeHelper.CreateService(Domain, UserName, Password, SenderEmail, Recipient, Subject, BodyFilePath, ExchangeServerUrl);

            Console.WriteLine("Creating message...");
            var message = new EmailMessage(service) { Subject = Subject, Body = new MessageBody(BodyType.HTML, body) };
            message.ToRecipients.Add(Recipient);

            Console.WriteLine("Sending message...");
            message.SendAndSaveCopy();

            return 0;
        }
    }
}
