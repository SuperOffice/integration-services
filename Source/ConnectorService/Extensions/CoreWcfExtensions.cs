using CoreWCF.Channels;
using CoreWCF;
using System.Runtime;
using CoreWCF.Configuration;
using CoreWCF.Description;
using ConnectorService.Services;
using SuperOffice.Online.IntegrationService.Contract;
using SuperOffice.SuperID.Contracts;
using Microsoft.Extensions.Options;
using ConnectorService.Models;
using SuperOffice.ErpSync;
using System.Web.Services.Description;
using static Org.BouncyCastle.Crypto.Engines.SM2Engine;
using Microsoft.Identity.Client;

namespace ConnectorService.Extensions
{
    public static class CoreWcfExtensions
    {
        /// <summary>
        /// Configures CoreWCF services and endpoints.
        /// </summary>
        internal static WebApplication AddWcfEndpoints(this WebApplication app)
        {
            var applicationOptions = app.GetApplicationOptions();
            app.UseServiceModel(serviceBuilder =>
            {
                serviceBuilder.AddCoreWcfServices(applicationOptions.Host);
                serviceBuilder.AddCoreWcfEndpoints();
            });
            return app;
        }

        internal static IServiceBuilder AddCoreWcfServices(this IServiceBuilder serviceBuilder, string host)
        {
            serviceBuilder.AddService<QuoteConnectorWS>(opts =>
            {
                opts.DebugBehavior.IncludeExceptionDetailInFaults = true;
                opts.DebugBehavior.HttpsHelpPageEnabled = true;
                opts.BaseAddresses.Add(new Uri($"http://{host}/Services/QuoteConnectorWS.svc"));
                opts.BaseAddresses.Add(new Uri($"https://{host}/Services/QuoteConnectorWS.svc"));
            });

            serviceBuilder.AddService<ErpConnectorWS>(opts =>
            {
                opts.DebugBehavior.IncludeExceptionDetailInFaults = true;
                opts.DebugBehavior.HttpsHelpPageEnabled = true;
                opts.BaseAddresses.Add(new Uri($"http://{host}/Services/ERPConnectorWS.svc"));
                opts.BaseAddresses.Add(new Uri($"https://{host}/Services/ERPConnectorWS.svc"));
            });
            return serviceBuilder;
        }

        internal static IServiceBuilder AddCoreWcfEndpoints(this IServiceBuilder serviceBuilder)
        {
            var quoteBinding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
            serviceBuilder.AddServiceEndpoint<QuoteConnectorWS, IOnlineQuoteConnector>(quoteBinding, string.Empty);
            serviceBuilder.AddServiceEndpoint<QuoteConnectorWS, IIntegrationServiceConnectorAuth>(quoteBinding, string.Empty);

            var erpBinding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
            serviceBuilder.AddServiceEndpoint<ErpConnectorWS, IErpConnectorWS>(erpBinding, string.Empty);
            serviceBuilder.AddServiceEndpoint<ErpConnectorWS, IIntegrationServiceConnectorAuth>(erpBinding, string.Empty);
            return serviceBuilder;
        }

        internal static Models.ApplicationOptions GetApplicationOptions(this WebApplication app)
        {
            var ioptions = app.Services.GetRequiredService<IOptions<Models.ApplicationOptions>>();
            return ioptions.Value;
        }

        internal static WebApplication EnableWsdlGet(this WebApplication app)
        {
            var serviceMetadataBehavior = app.Services.GetRequiredService<ServiceMetadataBehavior>();
            serviceMetadataBehavior.HttpGetEnabled = serviceMetadataBehavior.HttpsGetEnabled = true;
            return app;
        }
    }
}
