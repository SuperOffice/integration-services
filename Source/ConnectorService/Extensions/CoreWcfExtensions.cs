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

namespace ConnectorService.Extensions
{
    public static class CoreWcfExtensions
    {
        /// <summary>
        /// Configures CoreWCF services and endpoints.
        /// </summary>
        internal static WebApplication AddWcfEndpoints(this WebApplication app)
        {
            var wcfOptions = app.GetWcfOptions();

            app.UseServiceModel(serviceBuilder =>
            {
                serviceBuilder.AddCoreWcfServices();

                serviceBuilder.AddCoreWcfEndpoints(BasicHttpSecurityMode.None);

                if (wcfOptions.EnableHttpsEndpoints)
                {
                    serviceBuilder.AddCoreWcfEndpoints(BasicHttpSecurityMode.Transport);
                }
            });

            return app;
        }

        internal static IServiceBuilder AddCoreWcfServices(this IServiceBuilder serviceBuilder)
        {
            serviceBuilder.AddService<QuoteConnectorWS>();
            serviceBuilder.AddService<ErpConnectorWS>();
            return serviceBuilder;
        }

        internal static IServiceBuilder AddCoreWcfEndpoints(this IServiceBuilder serviceBuilder, BasicHttpSecurityMode mode)
        {
            var quoteBinding = new BasicHttpBinding(mode);
            serviceBuilder.AddServiceEndpoint<QuoteConnectorWS, IOnlineQuoteConnector>(quoteBinding, "Services/QuoteConnectorWS.svc");
            serviceBuilder.AddServiceEndpoint<QuoteConnectorWS, IIntegrationServiceConnectorAuth>(quoteBinding, "Services/QuoteConnectorWS.svc");

            var erpBinding = new BasicHttpBinding(mode);
            serviceBuilder.AddServiceEndpoint<ErpConnectorWS, IErpConnectorWS>(erpBinding, "Services/ErpConnectorWS.svc");
            serviceBuilder.AddServiceEndpoint<ErpConnectorWS, IIntegrationServiceConnectorAuth>(erpBinding, "Services/ErpConnectorWS.svc");
            return serviceBuilder;

        }

        internal static WcfOptions GetWcfOptions(this WebApplication app)
        {
            var ioptions = app.Services.GetRequiredService<IOptions<WcfOptions>>();
            return ioptions.Value;
        }

        internal static WebApplication EnableWsdlGet(this WebApplication app)
        {
            var serviceMetadataBehavior = app.Services.GetRequiredService<ServiceMetadataBehavior>();
            serviceMetadataBehavior.HttpGetEnabled = true;
            return app;
        }
    }
}
