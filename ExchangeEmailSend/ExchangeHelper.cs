
namespace ExchangeEmailSend
{
    using System;
    using System.Net;

    using Microsoft.Exchange.WebServices.Data;

    internal class ExchangeHelper
    {
        public static ExchangeService CreateService(string domain, string userName, string password, string senderEmail, string recipient, string subject, string bodyFilePath, string exchangeServerUrl = null)
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            var credential = new NetworkCredential(userName, password, domain);
            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1) { Credentials = credential };

            if (!string.IsNullOrEmpty(exchangeServerUrl))
            {
                service.Url = new Uri(exchangeServerUrl);
            }
            else
            {
                service.AutodiscoverUrl(senderEmail, url => url.Contains(domain));
            }

            return service;
        }
    }
}
